using DdddOcrSharp;
using Microsoft.Extensions.Logging;
using Microsoft.Playwright;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using System.Drawing;
using System.Globalization;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

namespace KillRiceMonkey.Infrastructure.Services;

public sealed class PlaywrightRuntime : IAsyncDisposable
{
    internal const string EmbeddedTemplatePrefix = "KillRiceMonkey.Infrastructure.TemplateImages.";
    internal const int PollDelayMilliseconds = 30;
    internal const int NolOcrScaleFactor = 3;
    private const int SmXvirtualscreen = 76;
    private const int SmYvirtualscreen = 77;
    private const int SmCxvirtualscreen = 78;
    private const int SmCyvirtualscreen = 79;
    private const int InputMouse = 0;
    private const uint MouseeventfLeftdown = 0x0002;
    private const uint MouseeventfLeftup = 0x0004;

    internal static readonly Regex StepPattern = new("^(?<step>\\d+)(?:-(?<suffix>[a-zA-Z0-9]+))?\\.png$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    internal static readonly Regex DigitsOnlyPattern = new("\\D", RegexOptions.Compiled);
    internal static readonly double[] MatchScales = [1.00, 0.95, 1.05, 0.90, 1.10];
    internal static readonly Dictionary<char, char> CaptchaCharCorrectionMap = new()
    {
        ['0'] = 'O', ['1'] = 'L', ['2'] = 'Z', ['3'] = 'E',
        ['4'] = 'A', ['5'] = 'S', ['6'] = 'D', ['7'] = 'T',
        ['8'] = 'B', ['9'] = 'Q',
        ['\u53EA'] = 'R', ['\u6C34'] = 'K', ['\u4E2D'] = 'P', ['\u5DF4'] = 'B',
        ['\u4E03'] = 'L', ['\u53E3'] = 'O', ['\u4E0A'] = 'W',
    };

    private static readonly HttpClient CdpHttpClient = new() { Timeout = TimeSpan.FromSeconds(2) };
    private static readonly HttpClient VisionApiClient = new() { Timeout = TimeSpan.FromSeconds(4) };
    private static Action<Exception, string>? _captchaWarningLogger;
    private static OcrEngine? _ocrEngine;
    private static DDDDOCR? _ddddOcrInstance;
    private static readonly object _ddddOcrLock = new();

    private readonly ILogger<PlaywrightRuntime> _logger;
    private IPlaywright? _playwright;

    public PlaywrightRuntime(ILogger<PlaywrightRuntime> logger)
    {
        _logger = logger;
        _captchaWarningLogger = (ex, message) => logger.LogWarning(ex, message);
    }

    public ValueTask DisposeAsync()
    {
        _playwright?.Dispose();
        _playwright = null;
        return ValueTask.CompletedTask;
    }

    internal async Task<IBrowser?> TryConnectToExistingChromiumBrowserAsync(string endpoint, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            _playwright ??= await Playwright.CreateAsync();
            return await _playwright.Chromium.ConnectOverCDPAsync(endpoint);
        }
        catch (PlaywrightException)
        {
            return null;
        }
    }

    internal static async Task<bool> IsCdpEndpointAvailableAsync(string endpoint, CancellationToken cancellationToken)
    {
        try
        {
            using var response = await CdpHttpClient.GetAsync(new Uri(new Uri(endpoint), "json/version"), cancellationToken);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    internal static bool IsClosedTargetError(PlaywrightException ex)
        => ex.Message.Contains("Target page, context or browser has been closed", StringComparison.OrdinalIgnoreCase)
        || ex.Message.Contains("Browser has been closed", StringComparison.OrdinalIgnoreCase)
        || ex.Message.Contains("Target closed", StringComparison.OrdinalIgnoreCase);

    internal static string NormalizeText(string? value)
        => string.IsNullOrWhiteSpace(value) ? string.Empty : Regex.Replace(value, "\\s+", " ").Trim();

    internal static string SafePageUrl(IPage page)
    {
        try { return page.Url; }
        catch (PlaywrightException ex) when (IsClosedTargetError(ex)) { return $"unavailable:{ex.Message}"; }
    }

    internal static string GetLogDirectoryPath() => Path.Combine(AppContext.BaseDirectory, "logs");

    internal static bool TryParseDesiredDate(string? value, out DateOnly date)
    {
        date = default;
        if (string.IsNullOrWhiteSpace(value)) return false;
        return DateOnly.TryParseExact(DigitsOnlyPattern.Replace(value, string.Empty), "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
    }

    internal static bool TryParseMonth(string value, out DateOnly month)
    {
        month = default;
        var digits = DigitsOnlyPattern.Replace(value, string.Empty);
        if (digits.Length != 6 || !int.TryParse(digits[..4], out var year) || !int.TryParse(digits[4..], out var monthValue) || monthValue is < 1 or > 12) return false;
        month = new DateOnly(year, monthValue, 1);
        return true;
    }

    internal static async Task WaitForConditionAsync(Func<Task<bool>> condition, TimeSpan timeout, CancellationToken cancellationToken, string errorMessage)
    {
        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (await condition()) return;
            await Task.Delay(PollDelayMilliseconds, cancellationToken);
        }
        throw new TimeoutException(errorMessage);
    }

    internal static async Task<bool> TryWaitForConditionAsync(Func<Task<bool>> condition, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (await condition()) return true;
            await Task.Delay(PollDelayMilliseconds, cancellationToken);
        }
        return false;
    }

    internal static async Task<string> GetPageTitleOrEmptyAsync(IPage page)
    {
        try { return await page.TitleAsync(); }
        catch (PlaywrightException ex) when (IsClosedTargetError(ex)) { return string.Empty; }
    }

    internal static async Task<string> GetLocatorTextOrEmptyAsync(ILocator locator)
    {
        try
        {
            if (await locator.CountAsync() == 0) return string.Empty;
            return NormalizeText(await locator.InnerTextAsync());
        }
        catch (PlaywrightException ex) when (IsClosedTargetError(ex))
        {
            return $"unavailable:{ex.Message}";
        }
    }

    internal static async Task ClickElementAsync(ILocator locator)
    {
        try { await locator.ClickAsync(new LocatorClickOptions { Force = true, Timeout = 1500 }); }
        catch (PlaywrightException) { await locator.EvaluateAsync("element => { element.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })); if (typeof element.click === 'function') { element.click(); } }"); }
    }

    internal static async Task ClickBookingButtonAsync(ILocator bookingButton, TimeSpan timeout)
    {
        var clickTimeout = (float)Math.Min(timeout.TotalMilliseconds, 5000);
        PlaywrightException? first = null;
        try { await bookingButton.ClickAsync(new LocatorClickOptions { Timeout = clickTimeout }); return; }
        catch (PlaywrightException ex) { first = ex; }
        try { await bookingButton.ClickAsync(new LocatorClickOptions { Force = true, Timeout = clickTimeout }); }
        catch (PlaywrightException ex) when (first is not null) { throw new InvalidOperationException($"예매하기 버튼 클릭에 실패했습니다. firstAttempt={first.Message}; secondAttempt={ex.Message}", ex); }
    }

    internal static ScreenFrame CaptureScreen()
    {
        var foregroundWindow = GetForegroundWindow();
        if (foregroundWindow != IntPtr.Zero && GetWindowRect(foregroundWindow, out var windowRect))
        {
            var foregroundWidth = windowRect.Right - windowRect.Left;
            var foregroundHeight = windowRect.Bottom - windowRect.Top;
            if (foregroundWidth > 100 && foregroundHeight > 100)
            {
                using var foregroundBitmap = new Bitmap(foregroundWidth, foregroundHeight);
                using (var foregroundGraphics = Graphics.FromImage(foregroundBitmap))
                {
                    foregroundGraphics.CopyFromScreen(windowRect.Left, windowRect.Top, 0, 0, new System.Drawing.Size(foregroundWidth, foregroundHeight));
                }
                return new ScreenFrame(BitmapConverter.ToMat(foregroundBitmap), windowRect.Left, windowRect.Top);
            }
        }

        var left = GetSystemMetrics(SmXvirtualscreen);
        var top = GetSystemMetrics(SmYvirtualscreen);
        var width = GetSystemMetrics(SmCxvirtualscreen);
        var height = GetSystemMetrics(SmCyvirtualscreen);
        using var bitmap = new Bitmap(width, height);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.CopyFromScreen(left, top, 0, 0, new System.Drawing.Size(width, height));
        }
        return new ScreenFrame(BitmapConverter.ToMat(bitmap), left, top);
    }

    internal static void ClickAt(int x, int y)
    {
        SetCursorPos(x, y);
        var inputs = new[]
        {
            new Input { Type = InputMouse, Union = new InputUnion { MouseInput = new MouseInput { DwFlags = MouseeventfLeftdown } } },
            new Input { Type = InputMouse, Union = new InputUnion { MouseInput = new MouseInput { DwFlags = MouseeventfLeftup } } },
        };
        SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<Input>());
    }

    internal static Mat ToGray(Mat source)
    {
        if (source.Channels() == 1) return source.Clone();
        var gray = new Mat();
        Cv2.CvtColor(source, gray, source.Channels() == 4 ? ColorConversionCodes.BGRA2GRAY : ColorConversionCodes.BGR2GRAY);
        return gray;
    }

    internal static IReadOnlyList<Mat> BuildScaledTemplates(Mat source)
    {
        var results = new List<Mat>();
        foreach (var scale in MatchScales)
        {
            if (Math.Abs(scale - 1.0) < 0.0001) { results.Add(source.Clone()); continue; }
            var width = (int)Math.Round(source.Width * scale);
            var height = (int)Math.Round(source.Height * scale);
            if (width < 2 || height < 2) continue;
            var resized = new Mat();
            Cv2.Resize(source, resized, new OpenCvSharp.Size(width, height), interpolation: InterpolationFlags.Linear);
            results.Add(resized);
        }
        return results;
    }

    internal static MatchHit? TryFindMatch(Mat grayScreenshot, IReadOnlyList<StepTemplate> templates, double threshold, out double bestScore, int? ignoreFromX = null)
    {
        bestScore = double.NegativeInfinity;
        foreach (var template in templates)
        {
            foreach (var scaledTemplate in template.ScaledTemplates)
            {
                if (grayScreenshot.Width < scaledTemplate.Width || grayScreenshot.Height < scaledTemplate.Height) continue;
                using var result = new Mat();
                Cv2.MatchTemplate(grayScreenshot, scaledTemplate, result, TemplateMatchModes.CCoeffNormed);
                Cv2.MinMaxLoc(result, out _, out var maxValue, out _, out var maxLocation);
                bestScore = Math.Max(bestScore, maxValue);
                var centerX = maxLocation.X + (scaledTemplate.Width / 2);
                if (ignoreFromX is not null && centerX >= ignoreFromX.Value) continue;
                if (maxValue >= threshold) return new MatchHit(template.State, centerX, maxLocation.Y + (scaledTemplate.Height / 2), maxValue);
            }
        }
        return null;
    }

    internal static TemplateBounds? TryFindTemplateBounds(Mat grayScreenshot, StepTemplate template, double threshold, out double bestScore)
    {
        bestScore = double.NegativeInfinity;
        foreach (var scaledTemplate in template.ScaledTemplates)
        {
            if (grayScreenshot.Width < scaledTemplate.Width || grayScreenshot.Height < scaledTemplate.Height) continue;
            using var result = new Mat();
            Cv2.MatchTemplate(grayScreenshot, scaledTemplate, result, TemplateMatchModes.CCoeffNormed);
            Cv2.MinMaxLoc(result, out _, out var maxValue, out _, out var maxLocation);
            bestScore = Math.Max(bestScore, maxValue);
            if (maxValue >= threshold) return new TemplateBounds(maxLocation.X, maxLocation.Y, scaledTemplate.Width, scaledTemplate.Height, maxValue);
        }
        return null;
    }

    internal static async Task<string> RecognizeTextAsync(Mat source, bool applyThreshold, CancellationToken cancellationToken)
    {
        using var prepared = PrepareOcrMat(source, applyThreshold);
        using var bgra = new Mat();
        if (prepared.Channels() == 4) prepared.CopyTo(bgra);
        else if (prepared.Channels() == 3) Cv2.CvtColor(prepared, bgra, ColorConversionCodes.BGR2BGRA);
        else Cv2.CvtColor(prepared, bgra, ColorConversionCodes.GRAY2BGRA);
        var size = checked((int)(bgra.Total() * bgra.ElemSize()));
        var bytes = new byte[size];
        Marshal.Copy(bgra.Data, bytes, 0, size);
        using var writer = new DataWriter();
        writer.WriteBytes(bytes);
        var buffer = writer.DetachBuffer();
        using var softwareBitmap = SoftwareBitmap.CreateCopyFromBuffer(buffer, BitmapPixelFormat.Bgra8, bgra.Width, bgra.Height, BitmapAlphaMode.Premultiplied);
        var result = await GetOcrEngine().RecognizeAsync(softwareBitmap);
        cancellationToken.ThrowIfCancellationRequested();
        return NormalizeText(result.Text);
    }

    internal static Mat PrepareOcrMat(Mat source, bool applyThreshold)
    {
        using var gray = ToGray(source);
        var resized = new Mat();
        Cv2.Resize(gray, resized, new OpenCvSharp.Size(gray.Width * NolOcrScaleFactor, gray.Height * NolOcrScaleFactor), interpolation: InterpolationFlags.Linear);
        if (!applyThreshold) return resized;
        var binary = new Mat();
        Cv2.GaussianBlur(resized, resized, new OpenCvSharp.Size(3, 3), 0);
        Cv2.Threshold(resized, binary, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);
        resized.Dispose();
        return binary;
    }

    internal static OcrEngine GetOcrEngine()
    {
        if (_ocrEngine is not null) return _ocrEngine;
        var preferred = new Language("ko-KR");
        if (OcrEngine.IsLanguageSupported(preferred)) _ocrEngine = OcrEngine.TryCreateFromLanguage(preferred);
        _ocrEngine ??= OcrEngine.TryCreateFromUserProfileLanguages();
        if (_ocrEngine is null)
        {
            var fallback = OcrEngine.AvailableRecognizerLanguages.FirstOrDefault();
            if (fallback is not null) _ocrEngine = OcrEngine.TryCreateFromLanguage(fallback);
        }
        return _ocrEngine ?? throw new InvalidOperationException("Windows OCR 엔진을 초기화하지 못했습니다. OCR 언어 팩 설치를 확인하세요.");
    }

    internal async Task<string> RecognizeCaptchaTextAsync(ILocator inputLocator, IPage page, IFrame? captchaFrame, CancellationToken cancellationToken)
    {
        if (page.IsClosed) return string.Empty;
        var imgLocator = await FindCaptchaImageAsync(inputLocator, page, captchaFrame);
        if (imgLocator is null) return string.Empty;
        byte[] screenshotBytes;
        try { screenshotBytes = await imgLocator.ScreenshotAsync(new LocatorScreenshotOptions { Timeout = 500 }); }
        catch (PlaywrightException) { return string.Empty; }
        var local = RunLocalCaptchaOcr(screenshotBytes);
        if (local.Length == 6) { _ = RecognizeCaptchaWithVisionApiAsync(screenshotBytes, cancellationToken); return local; }
        return await RecognizeCaptchaWithVisionApiAsync(screenshotBytes, cancellationToken);
    }

    internal string RunLocalCaptchaOcr(byte[] screenshotBytes)
    {
        var raw = RunDdddOcrOnBytes(screenshotBytes);
        return FilterCaptchaText(raw);
    }

    internal static string RunDdddOcrOnBytes(byte[] imageBytes)
    {
        try
        {
            if (_ddddOcrInstance is null)
            {
                lock (_ddddOcrLock)
                {
                    _ddddOcrInstance ??= new DDDDOCR(DdddOcrMode.ClassifyBeta);
                }
            }
            return _ddddOcrInstance.Classify(imageBytes) ?? string.Empty;
        }
        catch (Exception ex)
        {
            _captchaWarningLogger?.Invoke(ex, "Failed to initialize or run ddddocr for CAPTCHA");
            return string.Empty;
        }
    }

    internal async Task<string> RecognizeCaptchaWithVisionApiAsync(byte[] imageBytes, CancellationToken ct)
    {
        var apiKey = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
        if (string.IsNullOrWhiteSpace(apiKey)) return string.Empty;
        var requestBody = JsonSerializer.Serialize(new
        {
            model = "claude-sonnet-4-20250514",
            max_tokens = 32,
            messages = new[] { new { role = "user", content = new object[] { new { type = "image", source = new { type = "base64", media_type = "image/png", data = Convert.ToBase64String(imageBytes) } }, new { type = "text", text = "Read the 6 characters in this CAPTCHA image. Reply with ONLY the 6 uppercase characters, nothing else." } } } }
        });
        using var request = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages");
        request.Headers.Add("x-api-key", apiKey);
        request.Headers.Add("anthropic-version", "2023-06-01");
        request.Content = new StringContent(requestBody, System.Text.Encoding.UTF8, "application/json");
        try
        {
            using var response = await VisionApiClient.SendAsync(request, ct);
            if (!response.IsSuccessStatusCode) return string.Empty;
            using var doc = JsonDocument.Parse(await response.Content.ReadAsStringAsync(ct));
            foreach (var block in doc.RootElement.GetProperty("content").EnumerateArray())
            {
                if (block.GetProperty("type").GetString() == "text")
                {
                    var filtered = FilterCaptchaText(block.GetProperty("text").GetString() ?? string.Empty);
                    if (Regex.IsMatch(filtered, "^[A-Z0-9]{6}$")) return filtered;
                }
            }
        }
        catch { }
        return string.Empty;
    }

    internal string FilterCaptchaText(string raw)
    {
        var sb = new System.Text.StringBuilder(raw.Length);
        foreach (var ch in raw)
        {
            var upper = char.ToUpperInvariant(ch);
            if (CaptchaCharCorrectionMap.TryGetValue(upper, out var mapped)) sb.Append(mapped);
            else if (CaptchaCharCorrectionMap.TryGetValue(ch, out var mapped2)) sb.Append(mapped2);
            else if (char.IsAsciiLetterUpper(upper)) sb.Append(upper);
        }
        return sb.ToString();
    }

    internal static async Task<ILocator?> FindCaptchaImageAsync(ILocator inputLocator, IPage page, IFrame? captchaFrame)
    {
        for (var level = 1; level <= 2; level++)
        {
            var container = inputLocator.Locator($"xpath={string.Join("/", Enumerable.Repeat("..", level))}");
            var hinted = container.Locator("#imgCaptcha, #captchaImg, img[src*='captcha' i], img[src*='cap_img' i], [class*='captchaImage'] img");
            try { if (await hinted.CountAsync() > 0) return hinted.First; } catch { }
        }
        ILocator fp(string selector) => captchaFrame is not null ? captchaFrame.Locator(selector) : page.Locator(selector);
        var frameHinted = fp("#imgCaptcha, #captchaImg, img[src*='captcha' i], img[src*='cap_img' i], [class*='captchaImage'] img");
        try { if (await frameHinted.CountAsync() > 0) return frameHinted.First; } catch { }
        return null;
    }

    [DllImport("user32.dll")] private static extern int GetSystemMetrics(int nIndex);
    [DllImport("user32.dll")] private static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out Rect lpRect);
    [DllImport("user32.dll")] private static extern uint SendInput(uint nInputs, Input[] pInputs, int cbSize);

    [StructLayout(LayoutKind.Sequential)] internal struct Input { public int Type; public InputUnion Union; }
    [StructLayout(LayoutKind.Explicit)] internal struct InputUnion { [FieldOffset(0)] public MouseInput MouseInput; }
    [StructLayout(LayoutKind.Sequential)] internal struct MouseInput { public int Dx; public int Dy; public uint MouseData; public uint DwFlags; public uint Time; public nint DwExtraInfo; }
    [StructLayout(LayoutKind.Sequential)] internal struct Rect { public int Left; public int Top; public int Right; public int Bottom; }
    internal sealed class ScreenFrame : IDisposable { public ScreenFrame(Mat image, int offsetX, int offsetY) { Image = image; OffsetX = offsetX; OffsetY = offsetY; } public Mat Image { get; } public int OffsetX { get; } public int OffsetY { get; } public void Dispose() => Image.Dispose(); }
    internal readonly record struct TemplateBounds(int Left, int Top, int Width, int Height, double Score) { public int CenterX => Left + (Width / 2); public int CenterY => Top + (Height / 2); }
    internal readonly record struct MatchHit(string State, int X, int Y, double Score);
    internal readonly record struct PriorityCandidate(int X, int Y, double Score);
    internal sealed record TemplateMetadata(string State, int? Priority, bool IsViewMask);
    internal sealed record StepTemplate(string State, int? Priority, bool IsViewMask, IReadOnlyList<Mat> ScaledTemplates);
    internal sealed record StepGroup(int Step, IReadOnlyList<StepTemplate> Templates);
}
