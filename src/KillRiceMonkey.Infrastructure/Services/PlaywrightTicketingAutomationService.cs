using DdddOcrSharp;
using Microsoft.Playwright;
using Microsoft.Extensions.Logging;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Polly;
using System.Globalization;
using System.Drawing;
using System.Diagnostics;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using KillRiceMonkey.Application.Abstractions;
using KillRiceMonkey.Application.Models;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

namespace KillRiceMonkey.Infrastructure.Services;

public sealed partial class PlaywrightTicketingAutomationService : ITicketingAutomationService, IAsyncDisposable
{
    private const string EmbeddedTemplatePrefix = "KillRiceMonkey.Infrastructure.TemplateImages.";
    private static readonly Regex StepPattern = new("^(?<step>\\d+)(?:-(?<suffix>[a-zA-Z0-9]+))?\\.png$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex DigitsOnlyPattern = new("\\D", RegexOptions.Compiled);
    private static readonly Regex NolRoundPattern = new(@"^\D*(?<round>\d{1,2})\s*(?:회차|회|희|히|외)?\s*(?<time>\d{1,2}(?::|\.|,)?\d{2})", RegexOptions.Compiled);
    private static readonly double[] MatchScales = [1.00, 0.95, 1.05, 0.90, 1.10];
    private const int SeatSelectionOffset = 1;
    private const int Yes24LegendPaddingX = 8;
    private const double Yes24LegendSearchStartRatio = 0.70;
    private const double Yes24LegendMinIgnoreRatio = 0.55;
    private const double Yes24LegendThresholdDelta = 0.08;
    private const double Yes24LegendFallbackIgnoreRatio = 0.72;
    private const double Yes24LegendMaxIgnoreRatio = 0.92;
    private const int Yes24SeatColorSampleRadius = 6;
    private const double Yes24SeatMinSaturation = 35;

    private const int PollDelayMilliseconds = 30;

    private const int SmXvirtualscreen = 76;
    private const int SmYvirtualscreen = 77;
    private const int SmCxvirtualscreen = 78;
    private const int SmCyvirtualscreen = 79;
    private const int InputMouse = 0;
    private const uint MouseeventfLeftdown = 0x0002;
    private const uint MouseeventfLeftup = 0x0004;
    private const string NolRemoteDebugLaunchUrl = "https://tickets.interpark.com/";
    private const string NolCdpEndpoint = "http://127.0.0.1:9222/";
    private const string MelonRemoteDebugLaunchUrl = "https://ticket.melon.com/";
    private const int MelonRemoteDebugPort = 9223;
    private const string MelonCdpEndpoint = "http://127.0.0.1:9223/";
    private const string NolTemplateResourcePrefix = EmbeddedTemplatePrefix + "Nol.";
    private static readonly string[] NolBrowserExecutableCandidates =
    [
        @"C:\Program Files\Google\Chrome\Application\chrome.exe",
        @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        @"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    ];
    private const string NolPopupCloseTemplateFileName = "popup-close.png";
    private const string NolCalendarHeaderTemplateFileName = "calendar-header.png";
    private const string NolRoundHeaderTemplateFileName = "round-header.png";
    private const string NolBookingButtonTemplateFileName = "booking-button.png";
    private const double NolPanelToggleWidth = 290;
    private const double NolCalendarLeftOffset = 11;
    private const double NolCalendarTopOffset = 39;
    private const double NolCalendarWidth = 290;
    private const double NolCalendarHeight = 237;
    private const double NolCalendarMonthHeaderHeight = 28;
    private const double NolCalendarPrevArrowCenterX = 74;
    private const double NolCalendarNextArrowCenterX = 214;
    private const double NolCalendarArrowCenterY = 11;
    private const double NolCalendarGridLeftOffset = 14;
    private const double NolCalendarGridTopOffset = 66;
    private const double NolCalendarCellSize = 32;
    private const double NolCalendarColumnStep = 38;
    private const double NolCalendarRowStep = 34;
    private const double NolRoundListLeftOffset = 11;
    private const double NolRoundListTopOffset = 29;
    private const double NolRoundListWidth = 290;
    private const double NolRoundRowHeight = 45;
    private const int NolRoundMaxRows = 8;
    private const int NolMonthNavigationLimit = 24;
    private const int NolOcrScaleFactor = 3;

    private readonly ILogger<PlaywrightTicketingAutomationService> _logger;
    private readonly ResiliencePipeline _pipeline;
    private readonly SemaphoreSlim _nolBrowserLock = new(1, 1);
    private readonly SemaphoreSlim _melonBrowserLock = new(1, 1);
    private static readonly HttpClient NolHttpClient = new() { Timeout = TimeSpan.FromSeconds(2) };
    private static readonly HttpClient VisionApiClient = new() { Timeout = TimeSpan.FromSeconds(4) };
    private static Action<Exception, string>? _captchaWarningLogger;
    private static OcrEngine? _nolOcrEngine;
    private static DDDDOCR? _ddddOcrInstance;
    private static readonly object _ddddOcrLock = new();
    private IPlaywright? _playwright;
    private IBrowser? _preparedNolConnectedBrowser;
    private IPage? _preparedNolPage;
    private IBrowser? _preparedMelonConnectedBrowser;
    private IPage? _preparedMelonPage;
    private bool _popupClosedDuringPrepare;
    private bool _melonPopupClosedDuringPrepare;

    public PlaywrightTicketingAutomationService(ILogger<PlaywrightTicketingAutomationService> logger)
    {
        _logger = logger;
        _captchaWarningLogger = (ex, message) => logger.LogWarning(ex, message);
        _pipeline = new ResiliencePipelineBuilder()
            .AddRetry(new Polly.Retry.RetryStrategyOptions
            {
                MaxRetryAttempts = 2,
                Delay = TimeSpan.FromMilliseconds(300)
            })
            .Build();
    }

    public Task<AutomationRunResult> RunAsync(TicketingJobRequest request, CancellationToken cancellationToken)
        => RunAsync(request, null, cancellationToken);

    public async Task<AutomationRunResult> RunAsync(TicketingJobRequest request, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        try
        {
            return await _pipeline.ExecuteAsync(async token =>
            {
                if (request.TemplateType == TicketingTemplateType.Nol)
                {
                    return await RunNolAutomationAsync(request, progress, token);
                }

                if (request.TemplateType == TicketingTemplateType.Melon)
                {
                    return await RunMelonAutomationAsync(request, progress, token);
                }

                if (request.MatchThreshold <= 0 || request.MatchThreshold > 1)
                {
                    return new AutomationRunResult(false, "매칭 임계값은 0보다 크고 1 이하여야 합니다.", DateTimeOffset.Now);
                }

                if (request.StepTimeoutSeconds <= 0)
                {
                    return new AutomationRunResult(false, "단계별 제한 시간은 1초 이상이어야 합니다.", DateTimeOffset.Now);
                }

                var (stepGroups, loadError) = LoadStepGroups(request);
                if (loadError is not null)
                {
                    return new AutomationRunResult(false, loadError, DateTimeOffset.Now);
                }

                try
                {
                    if (stepGroups.Count == 0)
                    {
                        return new AutomationRunResult(false, "숫자.png 또는 숫자-상태.png 패턴의 이미지가 없습니다.", DateTimeOffset.Now);
                    }

                    _logger.LogInformation("Automation started. template={Template}, stepCount={Count}", request.TemplateType, stepGroups.Count);

                    foreach (var stepGroup in stepGroups)
                    {
                        token.ThrowIfCancellationRequested();

                        var found = await WaitAndClickStepAsync(stepGroup, request.TemplateType, request.MatchThreshold, request.StepTimeoutSeconds, token);
                        if (!found.IsSuccess)
                        {
                            return new AutomationRunResult(false, found.Message, DateTimeOffset.Now);
                        }
                    }

                    var message = $"{stepGroups.Count}개 단계를 완료했습니다.";
                    _logger.LogInformation(message);
                    return new AutomationRunResult(true, message, DateTimeOffset.Now);
                }
                finally
                {
                    DisposeStepGroups(stepGroups);
                }
            }, cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Automation failed with exception");
            return new AutomationRunResult(false, $"예외 발생: {ex.Message}", DateTimeOffset.Now);
        }
    }

    public async ValueTask DisposeAsync()
    {
        await ReleasePreparedNolConnectionAsync();
        await ReleasePreparedMelonConnectionAsync();
        _playwright?.Dispose();
        _playwright = null;
        _nolBrowserLock.Dispose();
        _melonBrowserLock.Dispose();
    }

    public async Task<bool> IsNolRemoteDebugBrowserAvailableAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return await IsNolCdpEndpointAvailableAsync(NolCdpEndpoint, cancellationToken);
    }

    public async Task<bool> IsNolAutomationPreparedAsync(CancellationToken cancellationToken)
    {
        return await TryGetPreparedNolPageAsync(cancellationToken) is not null;
    }

    public async Task<bool> IsMelonRemoteDebugBrowserAvailableAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return await IsNolCdpEndpointAvailableAsync(MelonCdpEndpoint, cancellationToken);
    }

    public async Task<bool> IsMelonAutomationPreparedAsync(CancellationToken cancellationToken)
    {
        return await TryGetPreparedMelonPageAsync(cancellationToken) is not null;
    }

    public async Task<string> LaunchNolRemoteDebugBrowserAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (await IsNolCdpEndpointAvailableAsync(NolCdpEndpoint, cancellationToken))
        {
            return $"이미 remote debug 브라우저가 열려 있습니다: {NolCdpEndpoint}";
        }

        var executablePath = NolBrowserExecutableCandidates.FirstOrDefault(File.Exists);
        if (string.IsNullOrWhiteSpace(executablePath))
        {
            throw new InvalidOperationException("Chrome 또는 Edge 실행 파일을 찾지 못했습니다.");
        }

        var profileDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "KillRiceMonkey",
            "NolRemoteDebugProfile");
        Directory.CreateDirectory(profileDirectory);

        var startInfo = new ProcessStartInfo
        {
            FileName = executablePath,
            Arguments = $"--remote-debugging-port=9222 --user-data-dir=\"{profileDirectory}\" --new-window \"{NolRemoteDebugLaunchUrl}\"",
            UseShellExecute = true
        };

        Process.Start(startInfo);

        var deadline = DateTimeOffset.UtcNow + TimeSpan.FromSeconds(10);
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (await IsNolCdpEndpointAvailableAsync(NolCdpEndpoint, cancellationToken))
            {
                return $"remote debug 브라우저를 열었습니다: {NolRemoteDebugLaunchUrl}";
            }

            await Task.Delay(200, cancellationToken);
        }

        return $"브라우저 실행은 요청했지만 remote debug 포트 확인이 지연되고 있습니다. 직접 확인: {NolCdpEndpoint}";
    }

    public async Task<string> LaunchMelonRemoteDebugBrowserAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (await IsNolCdpEndpointAvailableAsync(MelonCdpEndpoint, cancellationToken))
        {
            return $"이미 remote debug 브라우저가 열려 있습니다: {MelonCdpEndpoint}";
        }

        var executablePath = NolBrowserExecutableCandidates.FirstOrDefault(File.Exists);
        if (string.IsNullOrWhiteSpace(executablePath))
        {
            throw new InvalidOperationException("Chrome 또는 Edge 실행 파일을 찾지 못했습니다.");
        }

        var profileDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "KillRiceMonkey",
            "MelonRemoteDebugProfile");
        Directory.CreateDirectory(profileDirectory);

        var startInfo = new ProcessStartInfo
        {
            FileName = executablePath,
            Arguments = $"--remote-debugging-port={MelonRemoteDebugPort} --user-data-dir=\"{profileDirectory}\" --new-window \"{MelonRemoteDebugLaunchUrl}\"",
            UseShellExecute = true
        };

        Process.Start(startInfo);

        var deadline = DateTimeOffset.UtcNow + TimeSpan.FromSeconds(10);
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (await IsNolCdpEndpointAvailableAsync(MelonCdpEndpoint, cancellationToken))
            {
                return $"remote debug 브라우저를 열었습니다: {MelonRemoteDebugLaunchUrl}";
            }

            await Task.Delay(200, cancellationToken);
        }

        return $"브라우저 실행은 요청했지만 remote debug 포트 확인이 지연되고 있습니다. 직접 확인: {MelonCdpEndpoint}";
    }

    public async Task<string> PrepareNolAutomationAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!await IsNolRemoteDebugBrowserAvailableAsync(cancellationToken))
        {
            throw new InvalidOperationException("먼저 NOL Remote Debug 브라우저를 실행하세요.");
        }

        var page = await EnsurePreparedNolConnectedPageAsync(cancellationToken);
        if (page is null)
        {
            throw new InvalidOperationException("준비할 NOL 상품 페이지를 찾지 못했습니다. 상품 페이지를 연 뒤 다시 시도하세요.");
        }

        await page.BringToFrontAsync();
        await EnsureNolPopupClosedAsync(page, TimeSpan.FromSeconds(2), cancellationToken);
        _popupClosedDuringPrepare = true;
        var snapshot = await DescribeNolPageStateAsync(page);
        _logger.LogInformation("Prepared NOL automation state. {State}", snapshot);
        return $"NOL 준비 완료: {SafePageUrl(page)}";
    }

    public async Task<string> PrepareMelonAutomationAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!await IsMelonRemoteDebugBrowserAvailableAsync(cancellationToken))
        {
            throw new InvalidOperationException("먼저 Melon Remote Debug 브라우저를 실행하세요.");
        }

        var page = await EnsurePreparedMelonConnectedPageAsync(cancellationToken);
        if (page is null)
        {
            throw new InvalidOperationException("준비할 Melon 상품 페이지를 찾지 못했습니다. 상품 페이지를 연 뒤 다시 시도하세요.");
        }

        await page.BringToFrontAsync();
        await EnsureMelonPopupClosedAsync(page, TimeSpan.FromMilliseconds(500), cancellationToken);
        _melonPopupClosedDuringPrepare = true;
        return $"Melon 준비 완료: {SafePageUrl(page)}";
    }

    private static async Task<bool> IsNolCdpEndpointAvailableAsync(string endpoint, CancellationToken cancellationToken)
    {
        try
        {
            using var response = await NolHttpClient.GetAsync(new Uri(new Uri(endpoint), "json/version"), cancellationToken);
            return response.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    private static async Task<IBrowser?> TryConnectToExistingChromiumBrowserAsync(IPlaywright playwright, string endpoint, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            return await playwright.Chromium.ConnectOverCDPAsync(endpoint);
        }
        catch (PlaywrightException)
        {
            return null;
        }
    }



    private static Mat PrepareHsvCaptchaMask(Mat source)
    {
        var resized = new Mat();
        Cv2.Resize(source, resized, new OpenCvSharp.Size(source.Width * NolOcrScaleFactor, source.Height * NolOcrScaleFactor),
            interpolation: InterpolationFlags.Lanczos4);

        using var hsv = new Mat();
        Cv2.CvtColor(resized, hsv, ColorConversionCodes.BGR2HSV);
        var channels = Cv2.Split(hsv);
        using var vChannel = channels[2];
        channels[0].Dispose();
        channels[1].Dispose();

        var mask = new Mat();
        Cv2.Threshold(vChannel, mask, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);
        resized.Dispose();

        using var kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(3, 3));
        Cv2.MorphologyEx(mask, mask, MorphTypes.Close, kernel);

        ApplyConnectedComponentsFilter(mask, 300, 20000);

        return mask;
    }

    private static void ApplyConnectedComponentsFilter(Mat mask, int minArea, int maxArea)
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

    private static Mat PrepareKMeansCaptchaMask(Mat source)
    {
        var resized = new Mat();
        Cv2.Resize(source, resized, new OpenCvSharp.Size(source.Width * NolOcrScaleFactor, source.Height * NolOcrScaleFactor),
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

    private static Mat PrepareHsvAutoThresholdCaptchaMask(Mat source)
    {
        using var resized = new Mat();
        Cv2.Resize(source, resized, new OpenCvSharp.Size(source.Width * NolOcrScaleFactor, source.Height * NolOcrScaleFactor),
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

    private static Mat PrepareHsvSatValueCaptchaMask(Mat source)
    {
        using var resized = new Mat();
        Cv2.Resize(source, resized, new OpenCvSharp.Size(source.Width * NolOcrScaleFactor, source.Height * NolOcrScaleFactor),
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

    private static readonly Dictionary<char, char> CaptchaCharCorrectionMap = new()
    {
        ['0'] = 'O', ['1'] = 'L', ['2'] = 'Z', ['3'] = 'E',
        ['4'] = 'A', ['5'] = 'S', ['6'] = 'D', ['7'] = 'T',
        ['8'] = 'B', ['9'] = 'Q',
        ['\u53EA'] = 'R', ['\u6C34'] = 'K', ['\u4E2D'] = 'P', ['\u5DF4'] = 'B',
        ['\u4E03'] = 'L', ['\u53E3'] = 'O', ['\u4E0A'] = 'W',
    };

    private static string FilterCaptchaText(string raw)
    {
        var sb = new System.Text.StringBuilder(raw.Length);
        foreach (var ch in raw)
        {
            var upper = char.ToUpperInvariant(ch);
            if (CaptchaCharCorrectionMap.TryGetValue(upper, out var mapped))
            {
                sb.Append(mapped);
            }
            else if (CaptchaCharCorrectionMap.TryGetValue(ch, out var mapped2))
            {
                sb.Append(mapped2);
            }
            else if (char.IsAsciiLetterUpper(upper))
            {
                sb.Append(upper);
            }
        }
        return sb.ToString();
    }

    private async Task<string> RecognizeCaptchaTextAsync(ILocator inputLocator, IPage page, IFrame? captchaFrame, CancellationToken cancellationToken)
    {
        if (page.IsClosed)
            return string.Empty;

        var imgLocator = await FindCaptchaImageAsync(inputLocator, page, captchaFrame);
        if (imgLocator is null)
        {
            _logger.LogWarning("CAPTCHA image element not found near input");
            return string.Empty;
        }

        try
        {
            var imageReady = await imgLocator.EvaluateAsync<bool>(@"el => {
                if (el.tagName === 'IMG') return el.complete && el.naturalWidth > 0;
                if (el.tagName === 'CANVAS') return el.width > 0 && el.height > 0;
                return true;
            }");

            if (!imageReady)
            {
                _logger.LogInformation("[CAPTCHA] 이미지 미로드 — onload 대기 시작.");
                imageReady = await imgLocator.EvaluateAsync<bool>(@"el => new Promise(resolve => {
                    if (el.tagName !== 'IMG' || (el.complete && el.naturalWidth > 0)) { resolve(true); return; }
                    const timer = setTimeout(() => resolve(false), 2000);
                    el.addEventListener('load', () => { clearTimeout(timer); resolve(true); }, { once: true });
                    el.addEventListener('error', () => { clearTimeout(timer); resolve(false); }, { once: true });
                })");

                if (!imageReady)
                {
                    _logger.LogWarning("CAPTCHA image not loaded after onload wait");
                    return string.Empty;
                }
            }
        }
        catch (PlaywrightException ex)
        {
            _logger.LogWarning(ex, "CAPTCHA image load check failed");
        }

        byte[] screenshotBytes;
        try
        {
            screenshotBytes = await imgLocator.ScreenshotAsync(new LocatorScreenshotOptions { Timeout = 500 });
        }
        catch (PlaywrightException ex)
        {
            _logger.LogWarning(ex, "Failed to screenshot CAPTCHA image element");
            return string.Empty;
        }

        var localResult = RunLocalCaptchaOcr(screenshotBytes);
        if (localResult.Length == 6)
        {
            _logger.LogInformation("CAPTCHA solved via local OCR: {Text}", localResult);
            _ = RecognizeCaptchaWithVisionApiAsync(screenshotBytes, cancellationToken);
            return localResult;
        }

        try
        {
            var visionResult = await RecognizeCaptchaWithVisionApiAsync(screenshotBytes, cancellationToken);
            if (!string.IsNullOrEmpty(visionResult))
            {
                _logger.LogInformation("CAPTCHA solved via Vision API: {Text}", visionResult);
                return visionResult;
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Vision API CAPTCHA failed");
        }

        _logger.LogWarning("CAPTCHA OCR: all methods failed, returning empty");
        return string.Empty;
    }

    private string RunLocalCaptchaOcr(byte[] screenshotBytes)
    {
        var candidates = new System.Collections.Concurrent.ConcurrentBag<(string method, string filtered, double weight)>();

        var ddddocrRaw = RunDdddOcrOnBytes(screenshotBytes);
        var ddddocrFiltered = FilterCaptchaText(ddddocrRaw);
        if (ddddocrFiltered.Length == 6 && ddddocrFiltered.All(char.IsAsciiLetterOrDigit))
            candidates.Add(("original", ddddocrFiltered, 2.0));

        var variants = new (string name, int mode, double weight)[]
        {
            ("gray", 0, 1.0),
            ("binary", 1, 1.0),
            ("invert", 2, 1.5),
            ("contrast", 3, 1.0),
        };

        Parallel.ForEach(variants, item =>
        {
            try
            {
                using var source = Cv2.ImDecode(screenshotBytes, ImreadModes.Color);
                if (source.Empty()) return;
                using var processed = item.mode switch
                {
                    0 => PreprocessForDdddocr(source, DdddocrPreprocessMode.Grayscale),
                    1 => PreprocessForDdddocr(source, DdddocrPreprocessMode.Binary),
                    2 => PreprocessForDdddocr(source, DdddocrPreprocessMode.Invert),
                    3 => PreprocessForDdddocr(source, DdddocrPreprocessMode.Contrast),
                    _ => throw new ArgumentOutOfRangeException()
                };
                Cv2.ImEncode(".png", processed, out var pngBytes);
                var raw = RunDdddOcrOnBytes(pngBytes);
                var filtered = FilterCaptchaText(raw);
                if (filtered.Length == 6 && filtered.All(char.IsAsciiLetterOrDigit))
                    candidates.Add((item.name, filtered, item.weight));
            }
            catch { }
        });

        var validCandidates = candidates.ToList();
        if (validCandidates.Count == 0)
        {
            if (ddddocrFiltered.Length == 6)
                return ddddocrFiltered;
            return string.Empty;
        }

        var majorityLength = validCandidates
            .GroupBy(c => c.filtered.Length)
            .OrderByDescending(g => g.Sum(x => x.weight))
            .First().Key;
        var sameLenCandidates = validCandidates.Where(c => c.filtered.Length == majorityLength).ToList();

        var result = new char[majorityLength];
        for (var pos = 0; pos < majorityLength; pos++)
        {
            var votes = new Dictionary<char, double>();
            foreach (var c in sameLenCandidates)
            {
                var ch = c.filtered[pos];
                votes[ch] = votes.GetValueOrDefault(ch) + c.weight;
            }
            result[pos] = votes.OrderByDescending(v => v.Value).First().Key;
        }

        return new string(result);
    }

    private enum DdddocrPreprocessMode { Grayscale, Binary, Invert, Contrast }

    private static Mat PreprocessForDdddocr(Mat source, DdddocrPreprocessMode mode)
    {
        switch (mode)
        {
            case DdddocrPreprocessMode.Grayscale:
            {
                using var gray = new Mat();
                Cv2.CvtColor(source, gray, ColorConversionCodes.BGR2GRAY);
                var result = new Mat();
                Cv2.CvtColor(gray, result, ColorConversionCodes.GRAY2BGR);
                return result;
            }
            case DdddocrPreprocessMode.Binary:
            {
                using var gray = new Mat();
                Cv2.CvtColor(source, gray, ColorConversionCodes.BGR2GRAY);
                using var binary = new Mat();
                Cv2.Threshold(gray, binary, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);
                var result = new Mat();
                Cv2.CvtColor(binary, result, ColorConversionCodes.GRAY2BGR);
                return result;
            }
            case DdddocrPreprocessMode.Invert:
            {
                using var gray = new Mat();
                Cv2.CvtColor(source, gray, ColorConversionCodes.BGR2GRAY);
                using var inv = new Mat();
                Cv2.BitwiseNot(gray, inv);
                var result = new Mat();
                Cv2.CvtColor(inv, result, ColorConversionCodes.GRAY2BGR);
                return result;
            }
            case DdddocrPreprocessMode.Contrast:
            {
                using var lab = new Mat();
                Cv2.CvtColor(source, lab, ColorConversionCodes.BGR2Lab);
                var channels = Cv2.Split(lab);
                using var clahe = Cv2.CreateCLAHE(4.0, new OpenCvSharp.Size(8, 8));
                var enhanced = new Mat();
                clahe.Apply(channels[0], enhanced);
                channels[0].Dispose();
                channels[0] = enhanced;
                using var merged = new Mat();
                Cv2.Merge(channels, merged);
                foreach (var ch in channels) ch.Dispose();
                var result = new Mat();
                Cv2.CvtColor(merged, result, ColorConversionCodes.Lab2BGR);
                return result;
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(mode));
        }
    }

    private static (string text, float confidence) RunCaptchaOcrOnMat(Mat source)
    {
        var tessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
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

    private static string RunDdddOcrOnBytes(byte[] imageBytes)
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

    private async Task<string> RecognizeCaptchaWithVisionApiAsync(byte[] imageBytes, CancellationToken ct)
    {
        var apiKey = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");
        if (string.IsNullOrWhiteSpace(apiKey))
            return string.Empty;

        var base64Image = Convert.ToBase64String(imageBytes);

        const int maxAttempts = 2;
        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
                timeoutCts.CancelAfter(TimeSpan.FromSeconds(4));

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

                using var response = await VisionApiClient.SendAsync(request, timeoutCts.Token);
                if (!response.IsSuccessStatusCode)
                {
                    var statusCode = (int)response.StatusCode;
                    _logger.LogWarning("Vision API returned {StatusCode} (attempt {Attempt}/{Max})",
                        response.StatusCode, attempt, maxAttempts);
                    if (statusCode >= 500 && attempt < maxAttempts)
                        continue;
                    return string.Empty;
                }

                var responseJson = await response.Content.ReadAsStringAsync(timeoutCts.Token);
                using var doc = JsonDocument.Parse(responseJson);
                var contentArray = doc.RootElement.GetProperty("content");
                foreach (var block in contentArray.EnumerateArray())
                {
                    if (block.GetProperty("type").GetString() == "text")
                    {
                        var raw = block.GetProperty("text").GetString()?.Trim() ?? string.Empty;
                        var filtered = FilterCaptchaText(raw);
                        if (Regex.IsMatch(filtered, "^[A-Z0-9]{6}$"))
                        {
                            _logger.LogInformation("Vision API CAPTCHA result: {Text}", filtered);
                            return filtered;
                        }
                    }
                }

                _logger.LogWarning("Vision API returned non-6-char result (attempt {Attempt}/{Max})", attempt, maxAttempts);
                return string.Empty;
            }
            catch (TaskCanceledException ex) when (!ct.IsCancellationRequested)
            {
                _logger.LogWarning(ex, "Vision API timeout (attempt {Attempt}/{Max})", attempt, maxAttempts);
                if (attempt < maxAttempts)
                    continue;
            }
            catch (HttpRequestException ex)
            {
                _logger.LogWarning(ex, "Vision API request failed (attempt {Attempt}/{Max})", attempt, maxAttempts);
                if (attempt < maxAttempts)
                    continue;
            }
            catch (Exception ex) when (ex is JsonException)
            {
                _logger.LogWarning(ex, "Vision API response parse failed");
                return string.Empty;
            }
        }

        return string.Empty;
    }

    private static async Task<ILocator?> FindCaptchaImageAsync(ILocator inputLocator, IPage page, IFrame? captchaFrame)
    {
        for (var level = 1; level <= 2; level++)
        {
            var xpath = string.Join("/", Enumerable.Repeat("..", level));
            var container = inputLocator.Locator($"xpath={xpath}");

            var hinted = container.Locator("#imgCaptcha, #captchaImg, img[src*='captcha' i], img[src*='cap_img' i], [class*='captchaImage'] img");
            try { if (await hinted.CountAsync() > 0) return hinted.First; } catch (PlaywrightException) { }

            var canvas = container.Locator("canvas");
            try { if (await canvas.CountAsync() > 0) return canvas.First; } catch (PlaywrightException) { }

            var imgs = container.Locator("img");
            try
            {
                var count = await imgs.CountAsync();
                for (var i = 0; i < count; i++)
                {
                    var img = imgs.Nth(i);
                    var box = await img.BoundingBoxAsync();
                    if (box is not null && box.Width >= 60 && box.Width <= 400 && box.Height >= 20 && box.Height <= 120)
                        return img;
                }
            }
            catch (PlaywrightException) { }
        }

        ILocator FrameOrPage(string selector) =>
            captchaFrame is not null ? captchaFrame.Locator(selector) : page.Locator(selector);

        var frameHinted = FrameOrPage("#imgCaptcha, #captchaImg, img[src*='captcha' i], img[src*='cap_img' i], [class*='captchaImage'] img");
        try { if (await frameHinted.CountAsync() > 0) return frameHinted.First; } catch (PlaywrightException) { }

        var frameCanvas = FrameOrPage("canvas");
        try { if (await frameCanvas.CountAsync() > 0) return frameCanvas.First; } catch (PlaywrightException) { }

        var frameImgs = FrameOrPage("img");
        try
        {
            var count = await frameImgs.CountAsync();
            for (var i = 0; i < count; i++)
            {
                var img = frameImgs.Nth(i);
                var box = await img.BoundingBoxAsync();
                if (box is not null && box.Width >= 60 && box.Width <= 400 && box.Height >= 20 && box.Height <= 120)
                    return img;
            }
        }
        catch (PlaywrightException) { }

        return null;
    }

    private void SaveCaptchaDebugImage(byte[] imageBytes, string label)
    {
        try
        {
            var dir = Path.Combine(GetLogDirectoryPath(), "captcha-debug");
            Directory.CreateDirectory(dir);
            var fileName = $"{label}-{DateTimeOffset.Now:yyyyMMdd-HHmmss-fff}.png";
            File.WriteAllBytes(Path.Combine(dir, fileName), imageBytes);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Failed to save CAPTCHA debug image");
        }
    }

    private void SaveCaptchaDebugImage(Mat mat, string label)
    {
        try
        {
            var dir = Path.Combine(GetLogDirectoryPath(), "captcha-debug");
            Directory.CreateDirectory(dir);
            var fileName = $"{label}-{DateTimeOffset.Now:yyyyMMdd-HHmmss-fff}.png";
            Cv2.ImWrite(Path.Combine(dir, fileName), mat);
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "Failed to save CAPTCHA debug image");
        }
    }


    private static async Task<string> GetPageTitleOrEmptyAsync(IPage page)
    {
        try
        {
            return await page.TitleAsync();
        }
        catch (PlaywrightException ex) when (IsClosedTargetError(ex))
        {
            return string.Empty;
        }
    }

    private static async Task WaitForConditionAsync(
        Func<Task<bool>> condition,
        TimeSpan timeout,
        CancellationToken cancellationToken,
        string errorMessage)
    {
        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (await condition())
            {
                return;
            }

            await Task.Delay(PollDelayMilliseconds, cancellationToken);
        }

        throw new TimeoutException(errorMessage);
    }

    private static async Task<bool> TryWaitForConditionAsync(
        Func<Task<bool>> condition,
        TimeSpan timeout,
        CancellationToken cancellationToken)
    {
        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (await condition())
            {
                return true;
            }

            await Task.Delay(PollDelayMilliseconds, cancellationToken);
        }

        return false;
    }

    private static bool TryParseDesiredDate(string? value, out DateOnly date)
    {
        date = default;
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var digits = DigitsOnlyPattern.Replace(value, string.Empty);
        return DateOnly.TryParseExact(digits, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
    }

    private static bool TryParseMonth(string value, out DateOnly month)
    {
        month = default;
        var digits = DigitsOnlyPattern.Replace(value, string.Empty);
        if (digits.Length != 6 || !int.TryParse(digits[..4], out var year) || !int.TryParse(digits[4..], out var monthValue))
        {
            return false;
        }

        if (monthValue < 1 || monthValue > 12)
        {
            return false;
        }

        month = new DateOnly(year, monthValue, 1);
        return true;
    }




    private static async Task<string> GetLocatorTextOrEmptyAsync(ILocator locator)
    {
        try
        {
            if (await locator.CountAsync() == 0)
            {
                return string.Empty;
            }

            return NormalizeText(await locator.InnerTextAsync());
        }
        catch (PlaywrightException ex) when (IsClosedTargetError(ex))
        {
            return $"unavailable:{ex.Message}";
        }
    }

    private static string SafePageUrl(IPage page)
    {
        try
        {
            return page.Url;
        }
        catch (PlaywrightException ex) when (IsClosedTargetError(ex))
        {
            return $"unavailable:{ex.Message}";
        }
    }

    private static string NormalizeText(string? value)
    {
        return string.IsNullOrWhiteSpace(value)
            ? string.Empty
            : Regex.Replace(value, "\\s+", " ").Trim();
    }

    private static string GetLogDirectoryPath()
    {
        return Path.Combine(AppContext.BaseDirectory, "logs");
    }

    private async Task<(bool IsSuccess, string Message)> WaitAndClickStepAsync(
        StepGroup stepGroup,
        TicketingTemplateType templateType,
        double threshold,
        int timeoutSeconds,
        CancellationToken cancellationToken)
    {
        var started = DateTimeOffset.UtcNow;
        var activeTemplates = stepGroup.Templates.Where(x => x.State == "active").ToList();
        var normalTemplates = stepGroup.Templates.Where(x => x.State == "normal" && !x.IsViewMask).ToList();
        var singleTemplates = stepGroup.Templates.Where(x => x.State == "single" && !x.IsViewMask).ToList();
        var viewTemplates = stepGroup.Templates.Where(x => x.IsViewMask).ToList();
        var hasPriorityTemplates = singleTemplates.Any(x => x.Priority is not null);

        if (normalTemplates.Count == 0 && singleTemplates.Count == 0)
        {
            return (false, $"{stepGroup.Step}단계 클릭 가능한 이미지가 없습니다.");
        }

        var bestScore = double.NegativeInfinity;
        int? cachedIgnoreFromX = null;

        while (DateTimeOffset.UtcNow - started < TimeSpan.FromSeconds(timeoutSeconds))
        {
            cancellationToken.ThrowIfCancellationRequested();

            using var frame = CaptureScreen();
            using var grayFrame = ToGray(frame.Image);
            int? ignoreFromX = null;
            if (templateType == TicketingTemplateType.Yes24)
            {
                if (hasPriorityTemplates && normalTemplates.Count == 0)
                {
                    cachedIgnoreFromX ??= GetYes24FallbackIgnoreFromX(grayFrame.Width);
                }
                else
                {
                    cachedIgnoreFromX ??= ResolveYes24IgnoreFromX(grayFrame, viewTemplates, threshold);
                }

                ignoreFromX = cachedIgnoreFromX;
            }

            var clickableTemplates = normalTemplates.Count > 0 ? normalTemplates : singleTemplates;
            MatchHit? clickMatch;
            double clickBestScore;

            if (hasPriorityTemplates && normalTemplates.Count == 0)
            {
                if (templateType == TicketingTemplateType.Yes24 && ignoreFromX is not null && ignoreFromX.Value > 50 && ignoreFromX.Value < grayFrame.Width)
                {
                    using var seatRoiGray = new Mat(grayFrame, new OpenCvSharp.Rect(0, 0, ignoreFromX.Value, grayFrame.Height));
                    clickMatch = TryFindPrioritySeatMatch(frame.Image, seatRoiGray, clickableTemplates, threshold, null, true, out clickBestScore);
                }
                else
                {
                    clickMatch = TryFindPrioritySeatMatch(frame.Image, grayFrame, clickableTemplates, threshold, ignoreFromX, templateType == TicketingTemplateType.Yes24, out clickBestScore);
                }
            }
            else
            {
                clickMatch = TryFindMatch(grayFrame, clickableTemplates, threshold, out clickBestScore);
            }

            bestScore = Math.Max(bestScore, clickBestScore);
            if (clickMatch is not null)
            {
                ClickAt(frame.OffsetX + clickMatch.Value.X, frame.OffsetY + clickMatch.Value.Y);

                if (clickableTemplates == singleTemplates)
                {
                    return (true, $"{stepGroup.Step}단계 클릭 성공 (single, score={clickMatch.Value.Score:F3})");
                }

                var transitioned = await WaitForTransitionAsync(
                    activeTemplates,
                    normalTemplates,
                    threshold,
                    timeoutSeconds,
                    DateTimeOffset.UtcNow,
                    cancellationToken);

                if (transitioned)
                {
                    return (true, $"{stepGroup.Step}단계 클릭 성공 ({clickMatch.Value.State}, score={clickMatch.Value.Score:F3})");
                }

                return (false, $"{stepGroup.Step}단계 클릭 후 상태 전환 확인 실패");
            }

            await Task.Delay(PollDelayMilliseconds, cancellationToken);
        }

        var bestText = double.IsNegativeInfinity(bestScore) ? "N/A" : bestScore.ToString("F3");
        return (false, $"{stepGroup.Step}단계 버튼 탐지 실패 (timeout {timeoutSeconds}s, best={bestText})");
    }

    private static (IReadOnlyList<StepGroup> StepGroups, string? Error) LoadStepGroups(TicketingJobRequest request)
    {
        if (request.TemplateType == TicketingTemplateType.Custom)
        {
            var resolvedDirectory = ResolveImageDirectory(request.ImageDirectory);
            if (resolvedDirectory is null)
            {
                return (Array.Empty<StepGroup>(), $"이미지 폴더를 찾을 수 없습니다: {request.ImageDirectory}");
            }

            return (LoadStepGroupsFromDirectory(resolvedDirectory), null);
        }

        var embeddedResult = LoadStepGroupsFromEmbeddedResources(request.TemplateType);
        if (embeddedResult.Error is null)
        {
            return embeddedResult;
        }

        var resolvedFallbackDirectory = ResolveImageDirectory(request.ImageDirectory);
        if (resolvedFallbackDirectory is null)
        {
            return embeddedResult;
        }

        return (LoadStepGroupsFromDirectory(resolvedFallbackDirectory), null);
    }

    private static IReadOnlyList<StepGroup> LoadStepGroupsFromDirectory(string directory)
    {
        var map = new Dictionary<int, List<StepTemplate>>();

        foreach (var filePath in Directory.EnumerateFiles(directory, "*", SearchOption.TopDirectoryOnly))
        {
            var fileName = Path.GetFileName(filePath);
            var match = StepPattern.Match(fileName);
            if (!match.Success)
            {
                continue;
            }

            if (!int.TryParse(match.Groups["step"].Value, out var step))
            {
                continue;
            }

            var metadata = ParseTemplateMetadata(match);
            if (metadata is null)
            {
                continue;
            }

            var image = Cv2.ImRead(filePath, ImreadModes.Grayscale);
            if (image.Empty())
            {
                continue;
            }

            if (!map.TryGetValue(step, out var templates))
            {
                templates = new List<StepTemplate>();
                map[step] = templates;
            }

            if (metadata.Priority is null && templates.Any(x => x.State == metadata.State && x.Priority is null && x.IsViewMask == metadata.IsViewMask))
            {
                image.Dispose();
                throw new InvalidOperationException($"중복 상태 이미지가 존재합니다: {step}-{metadata.State}");
            }

            templates.Add(new StepTemplate(metadata.State, metadata.Priority, metadata.IsViewMask, BuildScaledTemplates(image)));
            image.Dispose();
        }

        return map
            .OrderBy(kvp => kvp.Key)
            .Select(kvp => new StepGroup(kvp.Key, kvp.Value))
            .ToList();
    }

    private static (IReadOnlyList<StepGroup> StepGroups, string? Error) LoadStepGroupsFromEmbeddedResources(TicketingTemplateType templateType)
    {
        var map = new Dictionary<int, List<StepTemplate>>();
        var assembly = typeof(PlaywrightTicketingAutomationService).Assembly;
        var prefix = $"{EmbeddedTemplatePrefix}{templateType}.";

        var resourceNames = assembly
            .GetManifestResourceNames()
            .Where(name => name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (resourceNames.Count == 0)
        {
            return (Array.Empty<StepGroup>(), $"{templateType} 템플릿 이미지 리소스를 찾을 수 없습니다.");
        }

        foreach (var resourceName in resourceNames)
        {
            var fileName = resourceName[prefix.Length..];
            var match = StepPattern.Match(fileName);
            if (!match.Success)
            {
                continue;
            }

            if (!int.TryParse(match.Groups["step"].Value, out var step))
            {
                continue;
            }

            using var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream is null)
            {
                continue;
            }

            using var memory = new MemoryStream();
            stream.CopyTo(memory);
            using var encoded = Cv2.ImDecode(memory.ToArray(), ImreadModes.Grayscale);
            if (encoded.Empty())
            {
                continue;
            }

            var metadata = ParseTemplateMetadata(match);
            if (metadata is null)
            {
                continue;
            }

            if (!map.TryGetValue(step, out var templates))
            {
                templates = new List<StepTemplate>();
                map[step] = templates;
            }

            if (metadata.Priority is null && templates.Any(x => x.State == metadata.State && x.Priority is null && x.IsViewMask == metadata.IsViewMask))
            {
                throw new InvalidOperationException($"중복 상태 이미지가 존재합니다: {step}-{metadata.State}");
            }

            templates.Add(new StepTemplate(metadata.State, metadata.Priority, metadata.IsViewMask, BuildScaledTemplates(encoded)));
        }

        var result = map
            .OrderBy(kvp => kvp.Key)
            .Select(kvp => new StepGroup(kvp.Key, kvp.Value))
            .ToList();

        return (result, null);
    }

    private static TemplateMetadata? ParseTemplateMetadata(Match match)
    {
        var suffixGroup = match.Groups["suffix"];
        if (!suffixGroup.Success)
        {
            return new TemplateMetadata("single", null, false);
        }

        var suffix = suffixGroup.Value.ToLowerInvariant();
        if (suffix is "normal" or "active")
        {
            return new TemplateMetadata(suffix, null, false);
        }

        if (suffix == "view")
        {
            return new TemplateMetadata("single", null, true);
        }

        if (int.TryParse(suffix, out var priority))
        {
            return new TemplateMetadata("single", priority, false);
        }

        return null;
    }

    private static int? DetectLegendIgnoreFromX(Mat grayScreenshot, IReadOnlyList<StepTemplate> viewTemplates, double threshold)
    {
        if (viewTemplates.Count == 0)
        {
            return null;
        }

        var minLegendX = (int)Math.Round(grayScreenshot.Width * Yes24LegendSearchStartRatio);
        var minAllowedIgnoreX = (int)Math.Round(grayScreenshot.Width * Yes24LegendMinIgnoreRatio);
        var legendThreshold = Math.Max(0.60, threshold - Yes24LegendThresholdDelta);
        var candidateXs = new List<int>();

        foreach (var template in viewTemplates)
        {
            foreach (var scaledTemplate in template.ScaledTemplates)
            {
                if (grayScreenshot.Width < scaledTemplate.Width || grayScreenshot.Height < scaledTemplate.Height)
                {
                    continue;
                }

                using var result = new Mat();
                Cv2.MatchTemplate(grayScreenshot, scaledTemplate, result, TemplateMatchModes.CCoeffNormed);

                for (var y = 0; y < result.Rows; y++)
                {
                    for (var x = 0; x < result.Cols; x++)
                    {
                        var score = result.At<float>(y, x);
                        if (score < legendThreshold)
                        {
                            continue;
                        }

                        var centerX = x + (scaledTemplate.Width / 2);
                        if (centerX < minLegendX)
                        {
                            continue;
                        }

                        candidateXs.Add(x);
                    }
                }
            }
        }

        if (candidateXs.Count == 0)
        {
            return null;
        }

        var ignoreFromX = Math.Max(0, candidateXs.Min() - Yes24LegendPaddingX);
        return ignoreFromX >= minAllowedIgnoreX ? ignoreFromX : null;
    }

    private static int ResolveYes24IgnoreFromX(Mat grayScreenshot, IReadOnlyList<StepTemplate> viewTemplates, double threshold)
    {
        var fallback = GetYes24FallbackIgnoreFromX(grayScreenshot.Width);
        var detected = DetectLegendIgnoreFromX(grayScreenshot, viewTemplates, threshold);
        if (detected is null)
        {
            return fallback;
        }

        var maxAllowed = (int)Math.Round(grayScreenshot.Width * Yes24LegendMaxIgnoreRatio);
        if (detected.Value > maxAllowed)
        {
            return fallback;
        }

        return detected.Value;
    }

    private static int GetYes24FallbackIgnoreFromX(int frameWidth)
    {
        return (int)Math.Round(frameWidth * Yes24LegendFallbackIgnoreRatio);
    }

    private static MatchHit? TryFindPrioritySeatMatch(
        Mat colorScreenshot,
        Mat grayScreenshot,
        IReadOnlyList<StepTemplate> templates,
        double threshold,
        int? ignoreFromX,
        bool enforceColoredSeat,
        out double bestScore)
    {
        bestScore = double.NegativeInfinity;
        var priorityGroups = templates
            .Where(x => x.Priority is not null)
            .GroupBy(x => x.Priority!.Value)
            .OrderBy(x => x.Key);

        foreach (var priorityGroup in priorityGroups)
        {
            var candidates = new List<PriorityCandidate>();

            CollectPriorityCandidates(grayScreenshot, priorityGroup, threshold, ignoreFromX, candidates, fastPathOnly: true, ref bestScore);
            if (candidates.Count == 0)
            {
                CollectPriorityCandidates(grayScreenshot, priorityGroup, threshold, ignoreFromX, candidates, fastPathOnly: false, ref bestScore);
            }

            if (candidates.Count == 0)
            {
                continue;
            }

            var ordered = DeduplicateCandidates(candidates)
                .OrderBy(x => x.Y)
                .ThenBy(x => x.X)
                .ToList();

            if (enforceColoredSeat)
            {
                ordered = ordered
                    .Where(x => IsLikelyColoredSeat(colorScreenshot, x.X, x.Y))
                    .ToList();
            }

            if (ordered.Count == 0)
            {
                continue;
            }

            var selectedIndex = Math.Min(SeatSelectionOffset, ordered.Count - 1);
            var selected = ordered[selectedIndex];
            return new MatchHit("single", selected.X, selected.Y, selected.Score);
        }

        return null;
    }

    private static void CollectPriorityCandidates(
        Mat grayScreenshot,
        IGrouping<int, StepTemplate> priorityGroup,
        double threshold,
        int? ignoreFromX,
        ICollection<PriorityCandidate> candidates,
        bool fastPathOnly,
        ref double bestScore)
    {
        foreach (var template in priorityGroup)
        {
            var scaledTemplates = fastPathOnly ? template.ScaledTemplates.Take(1) : template.ScaledTemplates;
            foreach (var scaledTemplate in scaledTemplates)
            {
                if (grayScreenshot.Width < scaledTemplate.Width || grayScreenshot.Height < scaledTemplate.Height)
                {
                    continue;
                }

                using var result = new Mat();
                Cv2.MatchTemplate(grayScreenshot, scaledTemplate, result, TemplateMatchModes.CCoeffNormed);
                Cv2.MinMaxLoc(result, out _, out var maxValue, out _, out _);
                bestScore = Math.Max(bestScore, maxValue);

                if (maxValue < threshold)
                {
                    continue;
                }

                CollectMatchCandidates(result, scaledTemplate.Width, scaledTemplate.Height, threshold, ignoreFromX, candidates);
            }
        }
    }

    private static void CollectMatchCandidates(
        Mat result,
        int templateWidth,
        int templateHeight,
        double threshold,
        int? ignoreFromX,
        ICollection<PriorityCandidate> output)
    {
        for (var y = 0; y < result.Rows; y++)
        {
            for (var x = 0; x < result.Cols; x++)
            {
                var score = result.At<float>(y, x);
                if (score < threshold)
                {
                    continue;
                }

                var centerX = x + (templateWidth / 2);
                if (ignoreFromX is not null && centerX >= ignoreFromX.Value)
                {
                    continue;
                }

                var centerY = y + (templateHeight / 2);
                output.Add(new PriorityCandidate(centerX, centerY, score));
            }
        }
    }

    private static IReadOnlyList<PriorityCandidate> DeduplicateCandidates(IReadOnlyList<PriorityCandidate> candidates)
    {
        const int mergeRadius = 10;
        var ordered = candidates.OrderByDescending(x => x.Score).ToList();
        var deduped = new List<PriorityCandidate>();

        foreach (var candidate in ordered)
        {
            var duplicated = deduped.Any(existing =>
                Math.Abs(existing.X - candidate.X) <= mergeRadius &&
                Math.Abs(existing.Y - candidate.Y) <= mergeRadius);
            if (duplicated)
            {
                continue;
            }

            deduped.Add(candidate);
        }

        return deduped;
    }

    private static bool IsLikelyColoredSeat(Mat colorScreenshot, int centerX, int centerY)
    {
        if (colorScreenshot.Empty())
        {
            return false;
        }

        var left = Math.Max(0, centerX - Yes24SeatColorSampleRadius);
        var top = Math.Max(0, centerY - Yes24SeatColorSampleRadius);
        var right = Math.Min(colorScreenshot.Width - 1, centerX + Yes24SeatColorSampleRadius);
        var bottom = Math.Min(colorScreenshot.Height - 1, centerY + Yes24SeatColorSampleRadius);
        var width = right - left + 1;
        var height = bottom - top + 1;
        if (width <= 1 || height <= 1)
        {
            return false;
        }

        using var roi = new Mat(colorScreenshot, new OpenCvSharp.Rect(left, top, width, height));
        using var bgr = new Mat();
        if (roi.Channels() == 4)
        {
            Cv2.CvtColor(roi, bgr, ColorConversionCodes.BGRA2BGR);
        }
        else if (roi.Channels() == 1)
        {
            Cv2.CvtColor(roi, bgr, ColorConversionCodes.GRAY2BGR);
        }
        else
        {
            roi.CopyTo(bgr);
        }

        using var hsv = new Mat();
        Cv2.CvtColor(bgr, hsv, ColorConversionCodes.BGR2HSV);
        var mean = Cv2.Mean(hsv);
        return mean.Val1 >= Yes24SeatMinSaturation;
    }

    private static MatchHit? TryFindMatch(
        Mat grayScreenshot,
        IReadOnlyList<StepTemplate> templates,
        double threshold,
        out double bestScore,
        int? ignoreFromX = null)
    {
        bestScore = double.NegativeInfinity;

        foreach (var template in templates)
        {
            foreach (var scaledTemplate in template.ScaledTemplates)
            {
                if (grayScreenshot.Width < scaledTemplate.Width || grayScreenshot.Height < scaledTemplate.Height)
                {
                    continue;
                }

                using var result = new Mat();
                Cv2.MatchTemplate(grayScreenshot, scaledTemplate, result, TemplateMatchModes.CCoeffNormed);
                Cv2.MinMaxLoc(result, out _, out var maxValue, out _, out var maxLocation);
                bestScore = Math.Max(bestScore, maxValue);
                var centerX = maxLocation.X + (scaledTemplate.Width / 2);
                if (ignoreFromX is not null && centerX >= ignoreFromX.Value)
                {
                    continue;
                }

                if (maxValue >= threshold)
                {
                    return new MatchHit(
                        template.State,
                        centerX,
                        maxLocation.Y + (scaledTemplate.Height / 2),
                        maxValue);
                }
            }
        }

        return null;
    }

    private static TemplateBounds? TryFindTemplateBounds(
        Mat grayScreenshot,
        StepTemplate template,
        double threshold,
        out double bestScore)
    {
        bestScore = double.NegativeInfinity;

        foreach (var scaledTemplate in template.ScaledTemplates)
        {
            if (grayScreenshot.Width < scaledTemplate.Width || grayScreenshot.Height < scaledTemplate.Height)
            {
                continue;
            }

            using var result = new Mat();
            Cv2.MatchTemplate(grayScreenshot, scaledTemplate, result, TemplateMatchModes.CCoeffNormed);
            Cv2.MinMaxLoc(result, out _, out var maxValue, out _, out var maxLocation);
            bestScore = Math.Max(bestScore, maxValue);
            if (maxValue < threshold)
            {
                continue;
            }

            return new TemplateBounds(maxLocation.X, maxLocation.Y, scaledTemplate.Width, scaledTemplate.Height, maxValue);
        }

        return null;
    }

    private static IReadOnlyList<Mat> BuildScaledTemplates(Mat source)
    {
        var results = new List<Mat>();
        foreach (var scale in MatchScales)
        {
            if (Math.Abs(scale - 1.0) < 0.0001)
            {
                results.Add(source.Clone());
                continue;
            }

            var width = (int)Math.Round(source.Width * scale);
            var height = (int)Math.Round(source.Height * scale);
            if (width < 2 || height < 2)
            {
                continue;
            }

            var resized = new Mat();
            Cv2.Resize(source, resized, new OpenCvSharp.Size(width, height), interpolation: InterpolationFlags.Linear);
            results.Add(resized);
        }

        return results;
    }

    private static Mat ToGray(Mat source)
    {
        if (source.Channels() == 1)
        {
            return source.Clone();
        }

        var gray = new Mat();
        if (source.Channels() == 4)
        {
            Cv2.CvtColor(source, gray, ColorConversionCodes.BGRA2GRAY);
            return gray;
        }

        Cv2.CvtColor(source, gray, ColorConversionCodes.BGR2GRAY);
        return gray;
    }

    private static async Task<bool> WaitForTransitionAsync(
        IReadOnlyList<StepTemplate> activeTemplates,
        IReadOnlyList<StepTemplate> normalTemplates,
        double threshold,
        int timeoutSeconds,
        DateTimeOffset stepStarted,
        CancellationToken cancellationToken)
    {
        while (DateTimeOffset.UtcNow - stepStarted < TimeSpan.FromSeconds(timeoutSeconds))
        {
            cancellationToken.ThrowIfCancellationRequested();

            using var frame = CaptureScreen();
            using var grayFrame = ToGray(frame.Image);

            if (activeTemplates.Count > 0 && TryFindMatch(grayFrame, activeTemplates, threshold, out _) is not null)
            {
                return true;
            }

            if (normalTemplates.Count > 0 && TryFindMatch(grayFrame, normalTemplates, threshold, out _) is null)
            {
                return true;
            }

            await Task.Delay(PollDelayMilliseconds, cancellationToken);
        }

        return false;
    }

    private static void DisposeStepGroups(IEnumerable<StepGroup> groups)
    {
        foreach (var group in groups)
        {
            foreach (var template in group.Templates)
            {
                foreach (var scaledTemplate in template.ScaledTemplates)
                {
                    scaledTemplate.Dispose();
                }
            }
        }
    }

    private static string? ResolveImageDirectory(string configuredPath)
    {
        var primary = Path.GetFullPath(configuredPath, AppContext.BaseDirectory);
        if (Directory.Exists(primary))
        {
            return primary;
        }

        var fallback = Path.GetFullPath(configuredPath, Environment.CurrentDirectory);
        if (Directory.Exists(fallback))
        {
            return fallback;
        }

        return null;
    }

    private static ScreenFrame CaptureScreen()
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

    private static void ClickAt(int x, int y)
    {
        SetCursorPos(x, y);
        var inputs = new[]
        {
            new Input
            {
                Type = InputMouse,
                Union = new InputUnion { MouseInput = new MouseInput { DwFlags = MouseeventfLeftdown } }
            },
            new Input
            {
                Type = InputMouse,
                Union = new InputUnion { MouseInput = new MouseInput { DwFlags = MouseeventfLeftup } }
            }
        };

        SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<Input>());
    }

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out Rect lpRect);

    [DllImport("user32.dll")]
    private static extern uint SendInput(uint nInputs, Input[] pInputs, int cbSize);

    [StructLayout(LayoutKind.Sequential)]
    private struct Input
    {
        public int Type;
        public InputUnion Union;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)]
        public MouseInput MouseInput;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MouseInput
    {
        public int Dx;
        public int Dy;
        public uint MouseData;
        public uint DwFlags;
        public uint Time;
        public nint DwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    private sealed class ScreenFrame : IDisposable
    {
        public ScreenFrame(Mat image, int offsetX, int offsetY)
        {
            Image = image;
            OffsetX = offsetX;
            OffsetY = offsetY;
        }

        public Mat Image { get; }
        public int OffsetX { get; }
        public int OffsetY { get; }

        public void Dispose()
        {
            Image.Dispose();
        }
    }

    private readonly record struct TemplateBounds(int Left, int Top, int Width, int Height, double Score)
    {
        public int CenterX => Left + (Width / 2);
        public int CenterY => Top + (Height / 2);
    }

    private readonly record struct MatchHit(string State, int X, int Y, double Score);
    private readonly record struct PriorityCandidate(int X, int Y, double Score);
    private sealed record TemplateMetadata(string State, int? Priority, bool IsViewMask);

    private sealed record StepTemplate(string State, int? Priority, bool IsViewMask, IReadOnlyList<Mat> ScaledTemplates);
    private sealed record StepGroup(int Step, IReadOnlyList<StepTemplate> Templates);
}
