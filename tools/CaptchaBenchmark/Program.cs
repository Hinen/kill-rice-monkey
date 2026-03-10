using System.Text.Json;
using System.Text.RegularExpressions;
using DdddOcrSharp;
using OpenCvSharp;

const int ScaleFactor = 3;

var captchaType = args.Length > 0 ? args[0].ToLowerInvariant() : "new";
if (captchaType is not ("new" or "old"))
{
    Console.Error.WriteLine("Usage: CaptchaBenchmark [new|old]");
    return 1;
}

var projectDir = AppDomain.CurrentDomain.BaseDirectory;
var repoRoot = Path.GetFullPath(Path.Combine(projectDir, "..", "..", "..", "..", ".."));
var samplesDir = Path.Combine(repoRoot, "captcha-samples", captchaType);
var resultsCsvPath = Path.Combine(samplesDir, "benchmark_results.csv");
var compareCsvPath = Path.Combine(samplesDir, "benchmark_results_compare.csv");
var groundTruthCsvPath = Path.Combine(samplesDir, "ground_truth.csv");
var tessDataPath = Path.Combine(projectDir, "tessdata");

Console.WriteLine($"CAPTCHA type: {captchaType.ToUpperInvariant()}");

if (!Directory.Exists(samplesDir))
{
    Console.Error.WriteLine($"Samples directory not found: {samplesDir}");
    return 1;
}

if (!File.Exists(Path.Combine(tessDataPath, "eng.traineddata")))
{
    Console.Error.WriteLine($"tessdata not found: {tessDataPath}");
    return 1;
}

var groundTruth = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
var useVisionApi = false;
var apiKey = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");

if (File.Exists(groundTruthCsvPath))
{
    Console.WriteLine($"Loading ground truth from: {groundTruthCsvPath}");
    foreach (var line in await File.ReadAllLinesAsync(groundTruthCsvPath))
    {
        var parts = line.Split(',', 2);
        if (parts.Length == 2 && parts[1].Length == 6)
            groundTruth[parts[0]] = parts[1];
    }
    Console.WriteLine($"Loaded {groundTruth.Count} ground truth labels");
}
else if (!string.IsNullOrWhiteSpace(apiKey))
{
    useVisionApi = true;
    Console.WriteLine("Using Vision API for ground truth labels");
}
else
{
    Console.Error.WriteLine("No ground truth source: ground_truth.csv not found and ANTHROPIC_API_KEY not set");
    return 1;
}

var files = Enumerable.Range(1, 110)
    .Select(i => Path.Combine(samplesDir, $"captcha_{i:D3}.png"))
    .Where(File.Exists)
    .OrderBy(f => f)
    .ToList();

Console.WriteLine($"Found {files.Count} captcha images in {samplesDir}");
Console.WriteLine($"tessdata path: {tessDataPath}");
Console.WriteLine();

using var httpClient = useVisionApi ? new HttpClient { Timeout = TimeSpan.FromSeconds(15) } : null;

using var ddddOcr = new DDDDOCR(DdddOcrMode.ClassifyBeta);

var results = new List<(string filename, string gt, string kmeans, string hsvAuto, string hsvOtsu, string hsvSatVal, string ddddocr)>();
var csvLines = new List<string> { "filename,ground_truth,kmeans_result,hsv_auto_result,hsv_otsu_result,hsv_satval_result,ddddocr_result" };

foreach (var (filePath, idx) in files.Select((f, i) => (f, i)))
{
    var filename = Path.GetFileName(filePath);
    Console.Write($"[{idx + 1}/{files.Count}] {filename} ... ");

    string gt;
    if (useVisionApi)
    {
        gt = await CallVisionApiAsync(httpClient!, apiKey!, filePath);
        await Task.Delay(500);
    }
    else
    {
        groundTruth.TryGetValue(filename, out var label);
        gt = label ?? string.Empty;
    }

    var kmeans = RunOcrWithMask(filePath, tessDataPath, PrepareKMeansCaptchaMask);
    var hsvAuto = RunOcrWithMask(filePath, tessDataPath, PrepareHsvCaptchaMask);
    var hsvOtsu = RunOcrWithMask(filePath, tessDataPath, PrepareHsvOtsuCaptchaMask);
    var hsvSatVal = RunOcrWithMask(filePath, tessDataPath, PrepareHsvSatValueCaptchaMask);
    var ddddocrResult = RunDdddOcr(filePath, ddddOcr);

    results.Add((filename, gt, kmeans, hsvAuto, hsvOtsu, hsvSatVal, ddddocrResult));
    csvLines.Add($"{filename},{gt},{kmeans},{hsvAuto},{hsvOtsu},{hsvSatVal},{ddddocrResult}");

    var km = MatchTag(gt, kmeans);
    var ha = MatchTag(gt, hsvAuto);
    var ho = MatchTag(gt, hsvOtsu);
    var hs = MatchTag(gt, hsvSatVal);
    var dd = MatchTag(gt, ddddocrResult);
    Console.WriteLine($"GT={gt,-8} KM={kmeans,-8}{km}  HA={hsvAuto,-8}{ha}  HO={hsvOtsu,-8}{ho}  HS={hsvSatVal,-8}{hs}  DD={ddddocrResult,-8}{dd}");
}

await File.WriteAllLinesAsync(compareCsvPath, csvLines);
Console.WriteLine($"\nResults saved to: {compareCsvPath}");

PrintSummary("K-means", results.Select(r => (r.gt, r.kmeans)).ToList());
PrintSummary("HSV Auto-Threshold", results.Select(r => (r.gt, r.hsvAuto)).ToList());
PrintSummary("HSV Otsu", results.Select(r => (r.gt, r.hsvOtsu)).ToList());
PrintSummary("HSV Sat+Value", results.Select(r => (r.gt, r.hsvSatVal)).ToList());
PrintSummary("ddddocr", results.Select(r => (r.gt, r.ddddocr)).ToList());

Console.WriteLine();
Console.WriteLine("=== ACCURACY COMPARISON ===");
Console.WriteLine($"{"Method",-22} {"Correct",8} {"Total",6} {"Accuracy",10}");
Console.WriteLine(new string('-', 48));
PrintAccuracyLine("K-means", results.Select(r => (r.gt, r.kmeans)).ToList());
PrintAccuracyLine("HSV Auto-Threshold", results.Select(r => (r.gt, r.hsvAuto)).ToList());
PrintAccuracyLine("HSV Otsu", results.Select(r => (r.gt, r.hsvOtsu)).ToList());
PrintAccuracyLine("HSV Sat+Value", results.Select(r => (r.gt, r.hsvSatVal)).ToList());
PrintAccuracyLine("ddddocr", results.Select(r => (r.gt, r.ddddocr)).ToList());

var ensembleResults = results.Select(r =>
{
    var candidates = new[] { r.kmeans, r.hsvAuto, r.hsvOtsu, r.hsvSatVal, r.ddddocr };
    var valid = candidates.Where(c => Regex.IsMatch(c, "^[A-Z0-9]{6}$")).ToList();
    if (valid.Count == 0) return (r.gt, result: string.Empty);
    var winner = valid.GroupBy(c => c).OrderByDescending(g => g.Count()).First().Key;
    return (r.gt, result: winner);
}).ToList();
PrintAccuracyLine("** ENSEMBLE **", ensembleResults);

var anyCorrectResults = results.Select(r =>
{
    var candidates = new[] { r.kmeans, r.hsvAuto, r.hsvOtsu, r.hsvSatVal, r.ddddocr };
    var anyCorrect = candidates.Any(c => string.Equals(c, r.gt, StringComparison.OrdinalIgnoreCase));
    return (r.gt, result: anyCorrect ? r.gt : string.Empty);
}).ToList();
PrintAccuracyLine("Any-Correct (max)", anyCorrectResults);

Console.WriteLine(new string('=', 48));

return 0;

// --- Helper functions ---

static string MatchTag(string gt, string result)
    => !string.IsNullOrEmpty(gt) && string.Equals(gt, result, StringComparison.OrdinalIgnoreCase) ? "OK" : "  ";

static void PrintAccuracyLine(string name, List<(string gt, string result)> pairs)
{
    var withBoth = pairs.Count(p => !string.IsNullOrEmpty(p.gt) && !string.IsNullOrEmpty(p.result));
    var correct = pairs.Count(p => !string.IsNullOrEmpty(p.gt) && string.Equals(p.gt, p.result, StringComparison.OrdinalIgnoreCase));
    var total = pairs.Count(p => !string.IsNullOrEmpty(p.gt));
    var accuracy = total > 0 ? (double)correct / total * 100 : 0;
    Console.WriteLine($"{name,-22} {correct,8} {total,6} {accuracy,9:F1}%");
}

static void PrintSummary(string name, List<(string gt, string result)> pairs)
{
    var total = pairs.Count(p => !string.IsNullOrEmpty(p.gt));
    var correct = pairs.Count(p => !string.IsNullOrEmpty(p.gt) && string.Equals(p.gt, p.result, StringComparison.OrdinalIgnoreCase));
    var empty = pairs.Count(p => string.IsNullOrEmpty(p.result));
    var accuracy = total > 0 ? (double)correct / total * 100 : 0;
    Console.WriteLine($"\n--- {name} ---");
    Console.WriteLine($"  Correct: {correct}/{total}  Accuracy: {accuracy:F1}%  Empty: {empty}");
}

static string RunOcrWithMask(string imagePath, string tessDataPath, Func<Mat, Mat> prepareMask)
{
    try
    {
        using var source = Cv2.ImRead(imagePath, ImreadModes.Color);
        if (source.Empty()) return string.Empty;

        using var mask = prepareMask(source);
        var (text, _) = RunCaptchaOcrOnMat(mask, tessDataPath);
        return FilterCaptchaText(text);
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"  OCR error: {ex.Message}");
        return string.Empty;
    }
}

static Mat PrepareKMeansCaptchaMask(Mat source)
{
    var resized = new Mat();
    Cv2.Resize(source, resized, new OpenCvSharp.Size(source.Width * ScaleFactor, source.Height * ScaleFactor),
        interpolation: InterpolationFlags.Lanczos4);

    var samples = new Mat();
    resized.ConvertTo(samples, MatType.CV_32FC3);
    samples = samples.Reshape(1, resized.Rows * resized.Cols);

    const int k = 3;
    var labels = new Mat();
    var centers = new Mat();
    Cv2.Kmeans(samples, k, labels,
        new TermCriteria(CriteriaTypes.Eps | CriteriaTypes.MaxIter, 100, 0.2),
        5, KMeansFlags.PpCenters, centers);
    samples.Dispose();

    var clusterAreas = new int[k];
    labels.GetArray(out int[] labelData);
    foreach (var l in labelData)
        clusterAreas[l]++;

    var centerBrightness = new double[k];
    for (var i = 0; i < k; i++)
    {
        var b = centers.At<float>(i, 0);
        var g = centers.At<float>(i, 1);
        var r = centers.At<float>(i, 2);
        centerBrightness[i] = Math.Sqrt(r * r + g * g + b * b);
    }

    var brightestCluster = Enumerable.Range(0, k).OrderByDescending(i => centerBrightness[i]).First();

    var sortedByArea = Enumerable.Range(0, k).OrderByDescending(i => clusterAreas[i]).ToArray();
    var secondLargestCluster = sortedByArea.Length > 1 ? sortedByArea[1] : sortedByArea[0];

    var textCluster = brightestCluster == secondLargestCluster
        ? brightestCluster
        : (centerBrightness[brightestCluster] > centerBrightness[secondLargestCluster] * 1.3
            ? brightestCluster
            : secondLargestCluster);

    var mask = new Mat(resized.Rows, resized.Cols, MatType.CV_8UC1, Scalar.Black);
    for (var i = 0; i < labelData.Length; i++)
    {
        if (labelData[i] == textCluster)
            mask.Set(i / resized.Cols, i % resized.Cols, (byte)255);
    }

    labels.Dispose();
    centers.Dispose();
    resized.Dispose();

    ApplyConnectedComponentsFilter(mask, 300, 20000);

    return mask;
}

static Mat PrepareHsvCaptchaMask(Mat source)
{
    using var resized = new Mat();
    Cv2.Resize(source, resized, new OpenCvSharp.Size(source.Width * ScaleFactor, source.Height * ScaleFactor),
        interpolation: InterpolationFlags.Lanczos4);

    using var hsv = new Mat();
    Cv2.CvtColor(resized, hsv, ColorConversionCodes.BGR2HSV);

    var channels = Cv2.Split(hsv);
    using var hChannel = channels[0];
    using var sChannel = channels[1];
    var vChannel = channels[2];

    int[] thresholds = [150, 160, 170, 180, 190, 200];
    var bestThreshold = 180;
    var bestDiff = int.MaxValue;

    foreach (var thresh in thresholds)
    {
        using var testMask = new Mat();
        Cv2.Threshold(vChannel, testMask, thresh, 255, ThresholdTypes.Binary);

        using var ccLabels = new Mat();
        using var ccStats = new Mat();
        using var ccCentroids = new Mat();
        var numLabels = Cv2.ConnectedComponentsWithStats(testMask, ccLabels, ccStats, ccCentroids);

        var validComponents = 0;
        for (var i = 1; i < numLabels; i++)
        {
            var area = ccStats.At<int>(i, 4);
            if (area >= 300 && area <= 20000)
                validComponents++;
        }

        var diff = Math.Abs(validComponents - 6);
        if (diff < bestDiff)
        {
            bestDiff = diff;
            bestThreshold = thresh;
        }
    }

    var mask = new Mat();
    Cv2.Threshold(vChannel, mask, bestThreshold, 255, ThresholdTypes.Binary);
    vChannel.Dispose();

    ApplyConnectedComponentsFilter(mask, 300, 20000);

    using var kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(3, 3));
    Cv2.MorphologyEx(mask, mask, MorphTypes.Close, kernel);

    return mask;
}

static Mat PrepareHsvOtsuCaptchaMask(Mat source)
{
    using var resized = new Mat();
    Cv2.Resize(source, resized, new OpenCvSharp.Size(source.Width * ScaleFactor, source.Height * ScaleFactor),
        interpolation: InterpolationFlags.Lanczos4);

    using var hsv = new Mat();
    Cv2.CvtColor(resized, hsv, ColorConversionCodes.BGR2HSV);

    var channels = Cv2.Split(hsv);
    using var hChannel = channels[0];
    using var sChannel = channels[1];
    using var vChannel = channels[2];

    var mask = new Mat();
    Cv2.Threshold(vChannel, mask, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);

    ApplyConnectedComponentsFilter(mask, 300, 20000);

    using var kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(3, 3));
    Cv2.MorphologyEx(mask, mask, MorphTypes.Close, kernel);

    return mask;
}

static Mat PrepareHsvSatValueCaptchaMask(Mat source)
{
    using var resized = new Mat();
    Cv2.Resize(source, resized, new OpenCvSharp.Size(source.Width * ScaleFactor, source.Height * ScaleFactor),
        interpolation: InterpolationFlags.Lanczos4);

    using var hsv = new Mat();
    Cv2.CvtColor(resized, hsv, ColorConversionCodes.BGR2HSV);

    var channels = Cv2.Split(hsv);
    using var hChannel = channels[0];
    using var sChannel = channels[1];
    using var vChannel = channels[2];

    using var sMask = new Mat();
    Cv2.Threshold(sChannel, sMask, 50, 255, ThresholdTypes.Binary);

    using var vMask = new Mat();
    Cv2.Threshold(vChannel, vMask, 160, 255, ThresholdTypes.Binary);

    var mask = new Mat();
    Cv2.BitwiseAnd(sMask, vMask, mask);

    ApplyConnectedComponentsFilter(mask, 300, 20000);

    using var kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(3, 3));
    Cv2.MorphologyEx(mask, mask, MorphTypes.Close, kernel);

    return mask;
}

static void ApplyConnectedComponentsFilter(Mat mask, int minArea, int maxArea)
{
    using var ccLabels = new Mat();
    using var ccStats = new Mat();
    using var ccCentroids = new Mat();
    var numLabels = Cv2.ConnectedComponentsWithStats(mask, ccLabels, ccStats, ccCentroids);

    for (var i = 1; i < numLabels; i++)
    {
        var area = ccStats.At<int>(i, 4);
        if (area >= minArea && area <= maxArea)
            continue;
        var left = ccStats.At<int>(i, 0);
        var top = ccStats.At<int>(i, 1);
        var w = ccStats.At<int>(i, 2);
        var h = ccStats.At<int>(i, 3);
        Cv2.Rectangle(mask, new OpenCvSharp.Rect(left, top, w, h), Scalar.Black, -1);
    }
}

static (string text, float confidence) RunCaptchaOcrOnMat(Mat source, string tessDataPath)
{
    Cv2.ImEncode(".png", source, out var pngBytes);

    var psmModes = new[]
    {
        TesseractOCR.Enums.PageSegMode.SingleLine,
        TesseractOCR.Enums.PageSegMode.SingleWord,
        TesseractOCR.Enums.PageSegMode.Auto,
    };

    var best = (text: string.Empty, confidence: 0f);

    foreach (var psm in psmModes)
    {
        using var engine = new TesseractOCR.Engine(tessDataPath, TesseractOCR.Enums.Language.English, TesseractOCR.Enums.EngineMode.Default);
        engine.SetVariable("tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789");
        engine.SetVariable("tessedit_load_system_dawg", "0");
        engine.SetVariable("tessedit_load_freq_dawg", "0");

        using var img = TesseractOCR.Pix.Image.LoadFromMemory(pngBytes);
        using var page = engine.Process(img, psm);
        var text = page.Text?.Trim() ?? string.Empty;
        var conf = page.MeanConfidence;

        if (text.Length > best.text.Length || (text.Length == best.text.Length && conf > best.confidence))
            best = (text, conf);

        if (best.text.Length >= 4 && conf > 0.5f)
            break;
    }

    return best;
}

static string RunDdddOcr(string imagePath, DDDDOCR ocr)
{
    try
    {
        var imageBytes = File.ReadAllBytes(imagePath);
        var raw = ocr.Classify(imageBytes);
        return FilterCaptchaText(raw);
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"  ddddocr error: {ex.Message}");
        return string.Empty;
    }
}

static string FilterCaptchaText(string raw)
    => new(raw.Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());

static async Task<string> CallVisionApiAsync(HttpClient httpClient, string apiKey, string imagePath)
{
    try
    {
        var imageBytes = await File.ReadAllBytesAsync(imagePath);
        var base64Image = Convert.ToBase64String(imageBytes);

        var requestBody = new
        {
            model = "claude-sonnet-4-20250514",
            max_tokens = 32,
            messages = new[]
            {
                new
                {
                    role = "user",
                    content = new object[]
                    {
                        new { type = "image", source = new { type = "base64", media_type = "image/png", data = base64Image } },
                        new { type = "text", text = "Read the 6 characters in this CAPTCHA image. Reply with ONLY the 6 uppercase characters, nothing else." }
                    }
                }
            }
        };

        var json = JsonSerializer.Serialize(requestBody);
        using var request = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages");
        request.Headers.Add("x-api-key", apiKey);
        request.Headers.Add("anthropic-version", "2023-06-01");
        request.Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

        using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        using var response = await httpClient.SendAsync(request, cts.Token);

        if (!response.IsSuccessStatusCode)
        {
            var errorBody = await response.Content.ReadAsStringAsync();
            Console.Error.WriteLine($"  Vision API {response.StatusCode}: {errorBody[..Math.Min(200, errorBody.Length)]}");
            return string.Empty;
        }

        var responseJson = await response.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(responseJson);
        var contentArray = doc.RootElement.GetProperty("content");
        foreach (var block in contentArray.EnumerateArray())
        {
            if (block.GetProperty("type").GetString() == "text")
            {
                var raw = block.GetProperty("text").GetString()?.Trim() ?? string.Empty;
                var filtered = FilterCaptchaText(raw);
                if (Regex.IsMatch(filtered, "^[A-Z0-9]{6}$"))
                    return filtered;
            }
        }
    }
    catch (Exception ex) when (ex is TaskCanceledException or HttpRequestException or JsonException)
    {
        Console.Error.WriteLine($"  Vision API error: {ex.Message}");
    }

    return string.Empty;
}
