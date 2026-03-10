using System.Text.Json;
using System.Text.RegularExpressions;
using OpenCvSharp;

const int ScaleFactor = 3;

var projectDir = AppDomain.CurrentDomain.BaseDirectory;
var repoRoot = Path.GetFullPath(Path.Combine(projectDir, "..", "..", "..", "..", ".."));
var samplesDir = Path.Combine(repoRoot, "captcha-samples", "new");
var resultsCsvPath = Path.Combine(samplesDir, "benchmark_results.csv");
var groundTruthCsvPath = Path.Combine(samplesDir, "ground_truth.csv");
var tessDataPath = Path.Combine(projectDir, "tessdata");

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

var results = new List<(string filename, string visionLabel, string ocrResult, bool match)>();
var csvLines = new List<string> { "filename,vision_label,kmeans_tesseract_result,match" };

foreach (var (filePath, idx) in files.Select((f, i) => (f, i)))
{
    var filename = Path.GetFileName(filePath);
    Console.Write($"[{idx + 1}/{files.Count}] {filename} ... ");

    string visionLabel;
    if (useVisionApi)
    {
        visionLabel = await CallVisionApiAsync(httpClient!, apiKey!, filePath);
        await Task.Delay(500);
    }
    else
    {
        groundTruth.TryGetValue(filename, out var gt);
        visionLabel = gt ?? string.Empty;
    }

    var ocrResult = RunKMeansTesseractOcr(filePath, tessDataPath);

    var match = !string.IsNullOrEmpty(visionLabel) && !string.IsNullOrEmpty(ocrResult)
        && string.Equals(visionLabel, ocrResult, StringComparison.OrdinalIgnoreCase);
    results.Add((filename, visionLabel, ocrResult, match));
    csvLines.Add($"{filename},{visionLabel},{ocrResult},{match}");

    Console.WriteLine($"Vision={visionLabel,-8} OCR={ocrResult,-8} {(match ? "OK" : "MISS")}");
}

await File.WriteAllLinesAsync(resultsCsvPath, csvLines);
Console.WriteLine($"\nResults saved to: {resultsCsvPath}");

var total = results.Count;
var correct = results.Count(r => r.match);
var visionEmpty = results.Count(r => string.IsNullOrEmpty(r.visionLabel));
var ocrEmpty = results.Count(r => string.IsNullOrEmpty(r.ocrResult));
var withBothLabels = results.Count(r => !string.IsNullOrEmpty(r.visionLabel) && !string.IsNullOrEmpty(r.ocrResult));
var accuracy = withBothLabels > 0 ? (double)correct / withBothLabels * 100 : 0;

Console.WriteLine();
Console.WriteLine("=== BENCHMARK SUMMARY ===");
Console.WriteLine($"Total images:        {total}");
Console.WriteLine($"Vision API empty:    {visionEmpty}");
Console.WriteLine($"OCR empty:           {ocrEmpty}");
Console.WriteLine($"Both have labels:    {withBothLabels}");
Console.WriteLine($"Exact matches:       {correct}");
Console.WriteLine($"Accuracy:            {accuracy:F1}%");
Console.WriteLine("=========================");

return 0;

// --- Functions ---

static string RunKMeansTesseractOcr(string imagePath, string tessDataPath)
{
    try
    {
        using var source = Cv2.ImRead(imagePath, ImreadModes.Color);
        if (source.Empty()) return string.Empty;

        using var mask = PrepareKMeansCaptchaMask(source);
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

    using var ccLabels = new Mat();
    using var ccStats = new Mat();
    using var ccCentroids = new Mat();
    var numLabels = Cv2.ConnectedComponentsWithStats(mask, ccLabels, ccStats, ccCentroids);

    for (var i = 1; i < numLabels; i++)
    {
        var area = ccStats.At<int>(i, 4);
        if (area >= 300 && area <= 20000)
            continue;
        var left = ccStats.At<int>(i, 0);
        var top = ccStats.At<int>(i, 1);
        var w = ccStats.At<int>(i, 2);
        var h = ccStats.At<int>(i, 3);
        Cv2.Rectangle(mask, new OpenCvSharp.Rect(left, top, w, h), Scalar.Black, -1);
    }

    return mask;
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
