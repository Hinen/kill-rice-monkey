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

public sealed class PlaywrightTicketingAutomationService : ITicketingAutomationService, IAsyncDisposable
{
    private const string EmbeddedTemplatePrefix = "KillRiceMonkey.Infrastructure.TemplateImages.";
    private static readonly Regex StepPattern = new("^(?<step>\\d+)(?:-(?<suffix>[a-zA-Z0-9]+))?\\.png$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex DigitsOnlyPattern = new("\\D", RegexOptions.Compiled);
    private static readonly Regex NolRoundPattern = new(@"^\D*(?<round>\d{1,2})\s*(?:회차|회|희|히|외)?\s*(?<time>\d{1,2}(?::|\.|,)?\d{2})", RegexOptions.Compiled);
    private static readonly double[] MatchScales = [1.00, 0.95, 1.05, 0.90, 1.10];
    private const int SeatSelectionOffset = 5;
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
    private static readonly HttpClient NolHttpClient = new() { Timeout = TimeSpan.FromSeconds(2) };
    private static readonly HttpClient VisionApiClient = new() { Timeout = TimeSpan.FromSeconds(4) };
    private static Action<Exception, string>? _captchaWarningLogger;
    private static OcrEngine? _nolOcrEngine;
    private IPlaywright? _playwright;
    private IBrowser? _preparedNolConnectedBrowser;
    private IPage? _preparedNolPage;
    private bool _popupClosedDuringPrepare;

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

    public async Task<AutomationRunResult> RunAsync(TicketingJobRequest request, CancellationToken cancellationToken)
    {
        try
        {
            return await _pipeline.ExecuteAsync(async token =>
            {
                if (request.TemplateType == TicketingTemplateType.Nol)
                {
                    return await RunNolAutomationAsync(request, token);
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
        _playwright?.Dispose();
        _playwright = null;
        _nolBrowserLock.Dispose();
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
            throw new InvalidOperationException("Edge 또는 Chrome 실행 파일을 찾지 못했습니다.");
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

    private async Task<AutomationRunResult> RunNolAutomationAsync(TicketingJobRequest request, CancellationToken cancellationToken)
    {
        if (request.MatchThreshold <= 0 || request.MatchThreshold > 1)
        {
            return new AutomationRunResult(false, "매칭 임계값은 0보다 크고 1 이하여야 합니다.", DateTimeOffset.Now);
        }

        if (request.StepTimeoutSeconds <= 0)
        {
            return new AutomationRunResult(false, "단계별 제한 시간은 1초 이상이어야 합니다.", DateTimeOffset.Now);
        }

        if (!TryParseDesiredDate(request.DesiredDate, out var desiredDate))
        {
            return new AutomationRunResult(false, "관람일 형식이 올바르지 않습니다. 예: 2026.04.11", DateTimeOffset.Now);
        }

        if (string.IsNullOrWhiteSpace(request.DesiredRound))
        {
            return new AutomationRunResult(false, "회차 값이 비어 있습니다. 예: 1회 19:00", DateTimeOffset.Now);
        }

        var desiredRound = NormalizeText(request.DesiredRound);
        var timeout = TimeSpan.FromSeconds(request.StepTimeoutSeconds);
        var runId = Guid.NewGuid().ToString("N");
        var phase = "screen-start";
        var threshold = request.MatchThreshold;

        void SetStage(string stage)
        {
            phase = stage;
            _logger.LogInformation(
                "NOL stage entered. runId={RunId}, stage={Stage}, date={Date}, round={Round}, logDir={LogDir}",
                runId,
                phase,
                desiredDate,
                desiredRound,
                GetLogDirectoryPath());
        }

        try
        {
            _logger.LogInformation("NOL screen automation started. runId={RunId}, date={Date}, round={Round}", runId, desiredDate, desiredRound);

            SetStage("attach-existing-browser");
            var cdpResult = await TryRunNolAutomationViaConnectedBrowserAsync(desiredDate, desiredRound, timeout, cancellationToken);
            if (cdpResult is not null)
            {
                return cdpResult;
            }

            SetStage("close-popup");
            await TryClickNolScreenTemplateAsync(NolPopupCloseTemplateFileName, Math.Min(threshold, 0.80), TimeSpan.FromSeconds(2), cancellationToken);

            SetStage("select-date");
            await SelectNolDateByScreenAsync(desiredDate, Math.Min(threshold, 0.82), timeout, cancellationToken);

            SetStage("select-round");
            var selectedRound = await SelectNolRoundByScreenAsync(desiredRound, Math.Min(threshold, 0.78), timeout, cancellationToken);

            SetStage("click-booking");
            await ClickNolScreenTemplateAsync(NolBookingButtonTemplateFileName, Math.Min(threshold, 0.82), timeout, cancellationToken, "예매하기 버튼");
            SetStage("completed");

            var message = $"NOL 화면 자동화 완료: {desiredDate:yyyy.MM.dd} / {selectedRound} 선택 후 예매하기 클릭 완료.";
            return new AutomationRunResult(true, message, DateTimeOffset.Now);
        }
        catch (TimeoutException ex)
        {
            _logger.LogError(ex, "NOL screen automation failed. runId={RunId}, phase={Phase}, date={Date}, round={Round}, logDir={LogDir}", runId, phase, desiredDate, desiredRound, GetLogDirectoryPath());
            return new AutomationRunResult(false, $"NOL 화면 자동화 시간 초과: {ex.Message} | 상세 로그: {GetLogDirectoryPath()}", DateTimeOffset.Now);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "NOL screen automation failed. runId={RunId}, phase={Phase}, date={Date}, round={Round}, logDir={LogDir}", runId, phase, desiredDate, desiredRound, GetLogDirectoryPath());
            return new AutomationRunResult(false, $"NOL 화면 자동화 예외: {ex.Message} | 상세 로그: {GetLogDirectoryPath()}", DateTimeOffset.Now);
        }
    }

    private async Task<AutomationRunResult?> TryRunNolAutomationViaConnectedBrowserAsync(DateOnly desiredDate, string desiredRound, TimeSpan timeout, CancellationToken cancellationToken)
    {
        IPage? page = null;
        try
        {
            page = await EnsurePreparedNolConnectedPageAsync(cancellationToken);
            if (page is null)
            {
                _logger.LogInformation("No existing Chromium browser with CDP endpoint was found. Falling back to screen automation.");
                return null;
            }

            _logger.LogInformation("Connected-browser NOL automation selected page. url={Url}", SafePageUrl(page));
            await page.BringToFrontAsync();
            if (!_popupClosedDuringPrepare)
            {
                await EnsureNolPopupClosedAsync(page, TimeSpan.FromSeconds(2), cancellationToken);
            }
            _popupClosedDuringPrepare = false;
            await SelectNolDateAsync(page, desiredDate, timeout, cancellationToken);
            await SelectNolRoundAsync(page, desiredRound, timeout, cancellationToken);
            var captchaPage = await ClickNolBookingAsync(page, timeout, cancellationToken);
            await SolveCaptchaAsync(captchaPage, timeout, cancellationToken);
            return new AutomationRunResult(true, $"NOL 기존 브라우저 DOM 자동화 완료: {desiredDate:yyyy.MM.dd} / {desiredRound} 선택, CAPTCHA 입력 완료.", DateTimeOffset.Now);
        }
        catch (Exception ex)
        {
            if (page is not null)
            {
                throw new InvalidOperationException($"NOL 기존 브라우저 DOM 자동화 실패: {ex.Message}", ex);
            }

            _logger.LogWarning(ex, "Connected-browser NOL automation failed. Falling back to screen automation.");
            return null;
        }
    }

    private async Task<IPage?> TryGetPreparedNolPageAsync(CancellationToken cancellationToken)
    {
        await _nolBrowserLock.WaitAsync(cancellationToken);
        try
        {
            if (await IsPreparedNolPageReusableAsync(_preparedNolPage, cancellationToken))
            {
                return _preparedNolPage;
            }

            if (_preparedNolConnectedBrowser is not null)
            {
                var refreshedPage = await FindConnectedNolPageAsync(_preparedNolConnectedBrowser, cancellationToken);
                if (await IsPreparedNolPageReusableAsync(refreshedPage, cancellationToken))
                {
                    _preparedNolPage = refreshedPage;
                    return refreshedPage;
                }
            }

            return null;
        }
        finally
        {
            _nolBrowserLock.Release();
        }
    }

    private async Task<IPage?> EnsurePreparedNolConnectedPageAsync(CancellationToken cancellationToken)
    {
        await _nolBrowserLock.WaitAsync(cancellationToken);
        try
        {
            if (await IsPreparedNolPageReusableAsync(_preparedNolPage, cancellationToken))
            {
                return _preparedNolPage;
            }

            if (_preparedNolConnectedBrowser is not null)
            {
                var existingPage = await FindConnectedNolPageAsync(_preparedNolConnectedBrowser, cancellationToken);
                if (await IsPreparedNolPageReusableAsync(existingPage, cancellationToken))
                {
                    _preparedNolPage = existingPage;
                    return existingPage;
                }
            }

            _playwright ??= await Playwright.CreateAsync();
            _preparedNolConnectedBrowser = await TryConnectToExistingChromiumBrowserAsync(_playwright, cancellationToken);
            if (_preparedNolConnectedBrowser is null)
            {
                _preparedNolPage = null;
                return null;
            }

            _preparedNolPage = await FindConnectedNolPageAsync(_preparedNolConnectedBrowser, cancellationToken);
            return _preparedNolPage;
        }
        finally
        {
            _nolBrowserLock.Release();
        }
    }

    private static async Task<bool> IsPreparedNolPageReusableAsync(IPage? page, CancellationToken cancellationToken)
    {
        if (page is null || page.IsClosed)
        {
            return false;
        }

        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            return SafePageUrl(page).Contains("tickets.interpark.com/goods/", StringComparison.OrdinalIgnoreCase) &&
                   await page.Locator("#productSide").CountAsync() > 0;
        }
        catch (PlaywrightException ex) when (IsClosedTargetError(ex))
        {
            return false;
        }
    }

    private Task ReleasePreparedNolConnectionAsync()
    {
        _preparedNolPage = null;
        _preparedNolConnectedBrowser = null;
        _popupClosedDuringPrepare = false;
        return Task.CompletedTask;
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

    private static async Task<IBrowser?> TryConnectToExistingChromiumBrowserAsync(IPlaywright playwright, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            return await playwright.Chromium.ConnectOverCDPAsync(NolCdpEndpoint);
        }
        catch (PlaywrightException)
        {
            return null;
        }
    }

    private static async Task<IPage?> FindConnectedNolPageAsync(IBrowser browser, CancellationToken cancellationToken)
    {
        var candidates = new List<(IPage Page, int Score)>();
        foreach (var page in browser.Contexts.SelectMany(x => x.Pages).Where(x => !x.IsClosed))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var url = SafePageUrl(page);
            var score = 0;
            if (url.Contains("tickets.interpark.com", StringComparison.OrdinalIgnoreCase))
            {
                score += 2;
            }

            if (url.Contains("/goods/", StringComparison.OrdinalIgnoreCase))
            {
                score += 4;
            }

            if (await TryWaitForConditionAsync(async () => await page.Locator("#productSide").CountAsync() > 0, TimeSpan.FromMilliseconds(400), cancellationToken))
            {
                score += 10;
            }

            try
            {
                if (await page.EvaluateAsync<bool>("() => document.visibilityState === 'visible'"))
                {
                    score += 3;
                }

                if (await page.EvaluateAsync<bool>("() => document.hasFocus()"))
                {
                    score += 3;
                }
            }
            catch (PlaywrightException ex) when (IsClosedTargetError(ex))
            {
                continue;
            }

            if (score > 0)
            {
                candidates.Add((page, score));
            }
        }

        return candidates
            .OrderByDescending(x => x.Score)
            .Select(x => x.Page)
            .FirstOrDefault();
    }

    private static async Task SelectNolDateByScreenAsync(DateOnly desiredDate, double threshold, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var template = LoadNolScreenTemplate(NolCalendarHeaderTemplateFileName);
        try
        {
            var deadline = DateTimeOffset.UtcNow + timeout;
            var navigationCount = 0;
            while (DateTimeOffset.UtcNow < deadline)
            {
                cancellationToken.ThrowIfCancellationRequested();

                using var frame = CaptureScreen();
                using var grayFrame = ToGray(frame.Image);
                var anchor = TryFindTemplateBounds(grayFrame, template, threshold, out _);
                if (anchor is null)
                {
                    await Task.Delay(PollDelayMilliseconds, cancellationToken);
                    continue;
                }

                var scale = GetNolTemplateScale(anchor.Value, NolPanelToggleWidth);
                var calendarRect = CreateNolRect(anchor.Value.Left + (int)Math.Round(NolCalendarLeftOffset * scale), anchor.Value.Top + (int)Math.Round(NolCalendarTopOffset * scale), (int)Math.Round(NolCalendarWidth * scale), (int)Math.Round(NolCalendarHeight * scale), frame.Image.Width, frame.Image.Height);
                var monthRect = CreateNolRect(calendarRect.Left, calendarRect.Top, calendarRect.Width, (int)Math.Round(NolCalendarMonthHeaderHeight * scale), frame.Image.Width, frame.Image.Height);
                var monthText = await ReadNolOcrTextAsync(frame.Image, monthRect, cancellationToken);
                if (!TryParseMonth(monthText, out var displayedMonth))
                {
                    await Task.Delay(PollDelayMilliseconds, cancellationToken);
                    continue;
                }

                var targetMonth = new DateOnly(desiredDate.Year, desiredDate.Month, 1);
                var monthDifference = GetMonthDifference(displayedMonth, targetMonth);
                if (monthDifference == 0)
                {
                    var cellCenter = GetNolCalendarCellCenter(calendarRect, desiredDate, scale);
                    ClickAt(frame.OffsetX + cellCenter.X, frame.OffsetY + cellCenter.Y);
                    await Task.Delay(200, cancellationToken);
                    return;
                }

                if (++navigationCount > NolMonthNavigationLimit)
                {
                    throw new TimeoutException($"관람일 {desiredDate:yyyy.MM.dd} 이 있는 달로 이동하지 못했습니다.");
                }

                var arrowX = monthDifference > 0
                    ? calendarRect.Left + (int)Math.Round(NolCalendarNextArrowCenterX * scale)
                    : calendarRect.Left + (int)Math.Round(NolCalendarPrevArrowCenterX * scale);
                var arrowY = calendarRect.Top + (int)Math.Round(NolCalendarArrowCenterY * scale);
                ClickAt(frame.OffsetX + arrowX, frame.OffsetY + arrowY);
                await Task.Delay(250, cancellationToken);
            }

            throw new TimeoutException($"관람일 {desiredDate:yyyy.MM.dd} 선택에 실패했습니다.");
        }
        finally
        {
            DisposeStepTemplate(template);
        }
    }

    private static async Task<string> SelectNolRoundByScreenAsync(string desiredRound, double threshold, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var roundHeaderTemplate = LoadNolScreenTemplate(NolRoundHeaderTemplateFileName);
        var bookingButtonTemplate = LoadNolScreenTemplate(NolBookingButtonTemplateFileName);
        var observedRoundTexts = new HashSet<string>(StringComparer.Ordinal);
        try
        {
            var deadline = DateTimeOffset.UtcNow + timeout;
            while (DateTimeOffset.UtcNow < deadline)
            {
                cancellationToken.ThrowIfCancellationRequested();

                using var frame = CaptureScreen();
                using var grayFrame = ToGray(frame.Image);
                var anchor = TryFindTemplateBounds(grayFrame, roundHeaderTemplate, threshold, out _);
                if (anchor is null)
                {
                    await Task.Delay(PollDelayMilliseconds, cancellationToken);
                    continue;
                }

                var scale = GetNolTemplateScale(anchor.Value, NolPanelToggleWidth);
                var bookingBounds = TryFindTemplateBounds(grayFrame, bookingButtonTemplate, Math.Min(0.82, threshold + 0.02), out _);
                var roundTop = anchor.Value.Top + (int)Math.Round(NolRoundListTopOffset * scale);
                var roundBottom = bookingBounds is not null
                    ? bookingBounds.Value.Top - (int)Math.Round(12 * scale)
                    : roundTop + (int)Math.Round(NolRoundRowHeight * NolRoundMaxRows * scale);
                var roundHeight = Math.Max((int)Math.Round(NolRoundRowHeight * scale), roundBottom - roundTop);
                var roundArea = CreateNolRect(anchor.Value.Left + (int)Math.Round(NolRoundListLeftOffset * scale), roundTop, (int)Math.Round(NolRoundListWidth * scale), roundHeight, frame.Image.Width, frame.Image.Height);
                var maxRows = Math.Max(1, (int)Math.Ceiling(roundArea.Height / Math.Max(1, NolRoundRowHeight * scale)));

                for (var rowIndex = 0; rowIndex < maxRows; rowIndex++)
                {
                    var rowRect = CreateNolRect(roundArea.Left, roundArea.Top + (int)Math.Round(rowIndex * NolRoundRowHeight * scale), roundArea.Width, (int)Math.Round(NolRoundRowHeight * scale), frame.Image.Width, frame.Image.Height);
                    if (rowRect.Width <= 0 || rowRect.Height <= 0)
                    {
                        continue;
                    }

                    var recognizedText = NormalizeNolRoundOcrText(await ReadNolOcrTextAsync(frame.Image, rowRect, cancellationToken));
                    if (string.IsNullOrWhiteSpace(recognizedText))
                    {
                        continue;
                    }

                    if (observedRoundTexts.Count < 8)
                    {
                        observedRoundTexts.Add(recognizedText);
                    }

                    if (!IsMatchingNolRound(recognizedText, desiredRound))
                    {
                        continue;
                    }

                    ClickAt(frame.OffsetX + rowRect.Left + (rowRect.Width / 2), frame.OffsetY + rowRect.Top + (rowRect.Height / 2));
                    await Task.Delay(200, cancellationToken);
                    return recognizedText;
                }

                if (observedRoundTexts.Count == 0 &&
                    TryClickNolRoundByOrder(frame.Image, frame.OffsetX, frame.OffsetY, roundArea, desiredRound, scale, out var fallbackRound))
                {
                    await Task.Delay(200, cancellationToken);
                    return fallbackRound;
                }

                if (observedRoundTexts.Count == 0 &&
                    TryClickNolRoundRowByIndex(frame.OffsetX, frame.OffsetY, roundArea, desiredRound, scale, maxRows, frame.Image.Width, frame.Image.Height, out fallbackRound))
                {
                    await Task.Delay(200, cancellationToken);
                    return fallbackRound;
                }

                await Task.Delay(PollDelayMilliseconds, cancellationToken);
            }

            var observedCandidates = observedRoundTexts.Count == 0
                ? "없음"
                : string.Join(" | ", observedRoundTexts);
            throw new TimeoutException($"NOL 화면에서 회차 {desiredRound} 를 찾지 못했습니다. OCR 후보: {observedCandidates}");
        }
        finally
        {
            DisposeStepTemplate(roundHeaderTemplate);
            DisposeStepTemplate(bookingButtonTemplate);
        }
    }

    private static async Task<string> ReadNolOcrTextAsync(Mat source, OpenCvSharp.Rect rect, CancellationToken cancellationToken)
    {
        if (rect.Width <= 0 || rect.Height <= 0)
        {
            return string.Empty;
        }

        using var roi = new Mat(source, rect);
        var text = await RecognizeNolTextAsync(roi, applyThreshold: false, cancellationToken);
        if (!string.IsNullOrWhiteSpace(text))
        {
            return text;
        }

        return await RecognizeNolTextAsync(roi, applyThreshold: true, cancellationToken);
    }

    private static async Task<string> RecognizeNolTextAsync(Mat source, bool applyThreshold, CancellationToken cancellationToken)
    {
        using var prepared = PrepareNolOcrMat(source, applyThreshold);
        using var bgra = new Mat();
        if (prepared.Channels() == 4)
        {
            prepared.CopyTo(bgra);
        }
        else if (prepared.Channels() == 3)
        {
            Cv2.CvtColor(prepared, bgra, ColorConversionCodes.BGR2BGRA);
        }
        else
        {
            Cv2.CvtColor(prepared, bgra, ColorConversionCodes.GRAY2BGRA);
        }

        var size = checked((int)(bgra.Total() * bgra.ElemSize()));
        var bytes = new byte[size];
        Marshal.Copy(bgra.Data, bytes, 0, size);

        using var writer = new DataWriter();
        writer.WriteBytes(bytes);
        var buffer = writer.DetachBuffer();
        using var softwareBitmap = SoftwareBitmap.CreateCopyFromBuffer(buffer, BitmapPixelFormat.Bgra8, bgra.Width, bgra.Height, BitmapAlphaMode.Premultiplied);
        var result = await GetNolOcrEngine().RecognizeAsync(softwareBitmap);
        cancellationToken.ThrowIfCancellationRequested();
        return NormalizeText(result.Text);
    }

    private static Mat PrepareNolOcrMat(Mat source, bool applyThreshold)
    {
        using var gray = ToGray(source);
        var resized = new Mat();
        Cv2.Resize(gray, resized, new OpenCvSharp.Size(gray.Width * NolOcrScaleFactor, gray.Height * NolOcrScaleFactor), interpolation: InterpolationFlags.Linear);
        if (!applyThreshold)
        {
            return resized;
        }

        var binary = new Mat();
        Cv2.GaussianBlur(resized, resized, new OpenCvSharp.Size(3, 3), 0);
        Cv2.Threshold(resized, binary, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);
        resized.Dispose();
        return binary;
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

    private static string FilterCaptchaText(string raw)
        => new(raw.Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());

    private async Task<string> RecognizeCaptchaTextAsync(ILocator inputLocator, IPage page, IFrame? captchaFrame, CancellationToken cancellationToken)
    {
        var imgLocator = await FindCaptchaImageAsync(inputLocator, page, captchaFrame);
        if (imgLocator is null)
        {
            _logger.LogWarning("CAPTCHA image element not found near input");
            return string.Empty;
        }

        try
        {
            var tagName = await imgLocator.EvaluateAsync<string>("el => el.tagName.toLowerCase()");
            var src = await imgLocator.EvaluateAsync<string>("el => el.src || el.currentSrc || ''");
            var box = await imgLocator.BoundingBoxAsync();
            _logger.LogInformation("CAPTCHA image found: tag={Tag}, src={Src}, width={W}, height={H}",
                tagName, src?.Length > 120 ? src[..120] : src, box?.Width, box?.Height);

            if (box is not null && (box.Width > 500 || box.Height > 200 || box.Width < 20 || box.Height < 10))
            {
                _logger.LogWarning("CAPTCHA image dimensions suspicious: {W}x{H}, likely wrong element", box.Width, box.Height);
            }
        }
        catch (PlaywrightException ex)
        {
            _logger.LogDebug(ex, "Failed to read CAPTCHA image attributes");
        }

        byte[] screenshotBytes;
        try
        {
            screenshotBytes = await imgLocator.ScreenshotAsync();
        }
        catch (PlaywrightException ex)
        {
            _logger.LogWarning(ex, "Failed to screenshot CAPTCHA image element");
            return string.Empty;
        }

        SaveCaptchaDebugImage(screenshotBytes, "captcha-raw");

        var visionTask = RecognizeCaptchaWithVisionApiAsync(screenshotBytes, cancellationToken);
        var localTask = Task.Run(() => RunLocalCaptchaOcr(screenshotBytes), cancellationToken);

        var visionResult = await visionTask;
        if (!string.IsNullOrEmpty(visionResult))
        {
            _logger.LogInformation("CAPTCHA solved via Vision API: {Text}", visionResult);
            return visionResult;
        }

        var localResult = await localTask;
        if (!string.IsNullOrEmpty(localResult))
        {
            _logger.LogInformation("CAPTCHA solved via local OCR: {Text}", localResult);
            return localResult;
        }

        _logger.LogWarning("CAPTCHA OCR: all methods failed, returning empty");
        return string.Empty;
    }

    private string RunLocalCaptchaOcr(byte[] screenshotBytes)
    {
        var ddddocrResult = RunDdddOcrOnBytes(screenshotBytes);
        var ddddocrFiltered = FilterCaptchaText(ddddocrResult);
        _logger.LogInformation("CAPTCHA OCR ddddocr: raw={Raw}, filtered={Filtered}", ddddocrResult, ddddocrFiltered);
        if (Regex.IsMatch(ddddocrFiltered, "^[A-Z0-9]{6}$"))
            return ddddocrFiltered;

        using var source = Cv2.ImDecode(screenshotBytes, ImreadModes.Color);
        if (source.Empty())
            return string.Empty;

        var preprocessors = new (string name, Func<Mat, Mat> fn)[]
        {
            ("hsv-otsu",   PrepareHsvCaptchaMask),
            ("hsv-auto",   PrepareHsvAutoThresholdCaptchaMask),
            ("hsv-satval", PrepareHsvSatValueCaptchaMask),
            ("kmeans",     PrepareKMeansCaptchaMask),
        };

        var candidates = new List<(string method, string filtered, float conf)>();

        foreach (var (name, fn) in preprocessors)
        {
            using var mask = fn(source);
            SaveCaptchaDebugImage(mask, $"captcha-{name}");
            var (text, conf) = RunCaptchaOcrOnMat(mask);
            var filtered = FilterCaptchaText(text);
            _logger.LogInformation("CAPTCHA OCR {Method}: ocrText={OcrText}, filtered={Filtered}, confidence={Conf:F3}",
                name, text, filtered, conf);
            candidates.Add((name, filtered, conf));
        }

        var validCandidates = candidates
            .Where(c => Regex.IsMatch(c.filtered, "^[A-Z0-9]{6}$"))
            .ToList();

        if (validCandidates.Count > 0)
        {
            var winner = validCandidates
                .GroupBy(c => c.filtered)
                .OrderByDescending(g => g.Count())
                .ThenByDescending(g => g.Max(c => c.conf))
                .First();
            _logger.LogInformation("CAPTCHA OCR ensemble: selected={Text} (votes={Votes})", winner.Key, winner.Count());
            return winner.Key;
        }

        var (rawText, rawConf) = RunCaptchaOcrOnMat(source);
        var rawFiltered = FilterCaptchaText(rawText);
        if (Regex.IsMatch(rawFiltered, "^[A-Z0-9]{6}$") && rawConf >= 0.3f)
            return rawFiltered;

        var allCandidates = candidates
            .Select(c => (filtered: c.filtered, conf: c.conf))
            .Append((filtered: rawFiltered, conf: rawConf));
        var bestFallback = allCandidates.OrderByDescending(c => c.conf).First().filtered;
        _logger.LogWarning("CAPTCHA OCR: no valid 6-char result from local, returning best={Text}", bestFallback);
        return bestFallback;
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
            using var ocr = new DDDDOCR(DdddOcrMode.ClassifyBeta);
            return ocr.Classify(imageBytes) ?? string.Empty;
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

            var hinted = container.Locator("img[src*='captcha' i], img[src*='cap_img' i]");
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

        var frameHinted = FrameOrPage("img[src*='captcha' i], img[src*='cap_img' i]");
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

    private async Task SolveCaptchaAsync(IPage page, TimeSpan timeout, CancellationToken cancellationToken)
    {
        const int maxAttempts = 3;
        const int minLength = 4;
        const int maxLength = 8;

        try
        {
            await page.WaitForLoadStateAsync(LoadState.Load, new PageWaitForLoadStateOptions
            {
                Timeout = 5000
            });
        }
        catch (TimeoutException) { }

        _logger.LogInformation("CAPTCHA 입력창 대기 시작 (무한 대기, F9로 취소). url={Url}", SafePageUrl(page));
        var (inputLocator, captchaFrame) = await FindCaptchaInputAsync(page, Timeout.InfiniteTimeSpan, cancellationToken);
        if (inputLocator is null)
        {
            foreach (var contextPage in page.Context.Pages.Where(p => p != page && !p.IsClosed))
            {
                _logger.LogInformation("CAPTCHA input not on main page, searching popup. url={Url}", SafePageUrl(contextPage));
                (inputLocator, captchaFrame) = await FindCaptchaInputAsync(contextPage, Timeout.InfiniteTimeSpan, cancellationToken);
                if (inputLocator is not null)
                {
                    page = contextPage;
                    break;
                }
            }
        }

        if (inputLocator is null)
        {
            _logger.LogWarning("CAPTCHA input not found (cancelled). url={Url}, frameCount={FrameCount}",
                SafePageUrl(page), page.Frames.Count);
            return;
        }

        _logger.LogInformation("CAPTCHA input found. url={Url}, inFrame={InFrame}",
            SafePageUrl(page), captchaFrame is not null);

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var text = await RecognizeCaptchaTextAsync(inputLocator, page, captchaFrame, cancellationToken);
            _logger.LogInformation("CAPTCHA attempt {Attempt}/{Max}: text={Text} (len={Len})",
                attempt, maxAttempts, text, text.Length);

            if (text.Length < minLength || text.Length > maxLength)
            {
                _logger.LogWarning("CAPTCHA text length {Len} out of range [{Min},{Max}], 새로고침 후 재시도",
                    text.Length, minLength, maxLength);
                if (attempt < maxAttempts)
                    await TryRefreshCaptchaImageAsync(page, captchaFrame, cancellationToken);
                await Task.Delay(200, cancellationToken);
                continue;
            }

            try
            {
                await inputLocator.First.FillAsync(string.Empty, new LocatorFillOptions { Timeout = 3000 });
                await inputLocator.First.PressSequentiallyAsync(text, new LocatorPressSequentiallyOptions { Delay = 10 });
            }
            catch (TimeoutException)
            {
                _logger.LogWarning("CAPTCHA input fill timed out (element not visible), using JS fallback");
                try
                {
                    await inputLocator.First.EvaluateAsync(@"(el, val) => {
                        el.value = val;
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        if (typeof jQuery !== 'undefined') { jQuery(el).val(val).trigger('input').trigger('change'); }
                    }", text);
                }
                catch (PlaywrightException ex)
                {
                    _logger.LogWarning(ex, "CAPTCHA JS fill fallback also failed");
                    continue;
                }
            }

            string filledValue;
            try
            {
                filledValue = await inputLocator.First.EvaluateAsync<string>("el => el.value || ''");
            }
            catch (PlaywrightException)
            {
                filledValue = string.Empty;
            }

            if (string.IsNullOrEmpty(filledValue))
            {
                _logger.LogWarning("CAPTCHA fill verification failed: input is empty after fill, retrying");
                await Task.Delay(100, cancellationToken);
                continue;
            }

            _logger.LogInformation("CAPTCHA fill verified: filled={Filled}", filledValue);

            const string submitSelector = "button:has-text('입력완료'), a:has-text('입력완료')";
            ILocator? submitLocator = null;
            var submitCount = 0;

            if (captchaFrame is not null)
            {
                var frameSubmit = captchaFrame.Locator(submitSelector);
                try { submitCount = await frameSubmit.CountAsync(); } catch (PlaywrightException) { }
                if (submitCount > 0) submitLocator = frameSubmit;
            }

            if (submitLocator is null)
            {
                var pageSubmit = page.Locator(submitSelector);
                try
                {
                    var pageCount = await pageSubmit.CountAsync();
                    if (pageCount > 0)
                    {
                        submitLocator = pageSubmit;
                        submitCount = pageCount;
                    }
                }
                catch (PlaywrightException) { }
            }

            _logger.LogInformation("CAPTCHA submit search: count={Count}, inFrame={InFrame}", submitCount, captchaFrame is not null);

            if (submitLocator is not null && submitCount > 0)
            {
                try
                {
                    await submitLocator.Last.ClickAsync(new LocatorClickOptions { Timeout = 3000, Force = true });
                }
                catch (TimeoutException)
                {
                    _logger.LogWarning("CAPTCHA submit click timed out, using JS fallback");
                    try
                    {
                        await submitLocator.Last.EvaluateAsync(@"el => {
                            if (el.disabled) el.disabled = false;
                            el.click();
                        }");
                    }
                    catch (PlaywrightException)
                    {
                        try
                        {
                            await inputLocator.First.EvaluateAsync(@"el => {
                                if (typeof fnCheck === 'function') { fnCheck(); }
                                else if (el.form) { el.form.submit(); }
                                else { el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, bubbles: true })); }
                            }");
                        }
                        catch (PlaywrightException) { await inputLocator.First.PressAsync("Enter"); }
                    }
                }
            }
            else
            {
                _logger.LogWarning("CAPTCHA submit button not found, using JS fnCheck/Enter fallback");
                try
                {
                    await inputLocator.First.EvaluateAsync(@"el => {
                        if (typeof fnCheck === 'function') { fnCheck(); }
                        else if (el.form) { el.form.submit(); }
                        else { el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, bubbles: true })); }
                    }");
                }
                catch (PlaywrightException) { await inputLocator.First.PressAsync("Enter"); }
            }

            var submitted = await TryWaitForConditionAsync(
                async () =>
                {
                    try { return await inputLocator.CountAsync() == 0; }
                    catch (PlaywrightException) { return false; }
                },
                TimeSpan.FromSeconds(3),
                cancellationToken);

            if (submitted)
            {
                _logger.LogInformation("CAPTCHA solved on attempt {Attempt}", attempt);
                return;
            }

            _logger.LogWarning("CAPTCHA 인식 실패 (attempt {Attempt}/{Max}), 새로고침 시도", attempt, maxAttempts);
            if (attempt < maxAttempts)
            {
                await TryRefreshCaptchaImageAsync(page, captchaFrame, cancellationToken);
                await Task.Delay(200, cancellationToken);
            }
        }

        throw new InvalidOperationException($"CAPTCHA 자동 인식 실패 ({maxAttempts}회 시도).");
    }

    private async Task TryRefreshCaptchaImageAsync(IPage page, IFrame? captchaFrame, CancellationToken cancellationToken)
    {
        const string refreshSelector =
            "img[src*='reload' i], img[src*='refresh' i], img[src*='btn_re' i], img[alt*='새로' i], img[alt*='변경' i], img[alt*='refresh' i], " +
            "button:has-text('변경'), button:has-text('새로고침'), a:has-text('변경'), a:has-text('새로고침'), " +
            "button[onclick*='captcha' i], a[onclick*='captcha' i], " +
            "[class*='refresh' i], [class*='reload' i], [class*='btn_re' i]";

        ILocator FrameOrPage(string selector) =>
            captchaFrame is not null ? captchaFrame.Locator(selector) : page.Locator(selector);

        try
        {
            var refreshLocator = FrameOrPage(refreshSelector);
            var count = await refreshLocator.CountAsync();
            if (count > 0)
            {
                await refreshLocator.First.ClickAsync(new LocatorClickOptions { Timeout = 2000, Force = true });
                _logger.LogInformation("CAPTCHA 새로고침 버튼 클릭 완료");
                return;
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "CAPTCHA 새로고침 버튼 클릭 실패, JS fallback 시도");
        }

        try
        {
            var jsResult = await (captchaFrame is not null
                ? captchaFrame.EvaluateAsync<bool>(@"() => {
                    if (typeof fnRefresh === 'function') { fnRefresh(); return true; }
                    if (typeof captchaRefresh === 'function') { captchaRefresh(); return true; }
                    if (typeof refreshCaptcha === 'function') { refreshCaptcha(); return true; }
                    var imgs = document.querySelectorAll('img[src*=""captcha"" i], img[src*=""cap_img"" i]');
                    for (var img of imgs) { img.src = img.src.split('?')[0] + '?t=' + Date.now(); return true; }
                    return false;
                }")
                : page.EvaluateAsync<bool>(@"() => {
                    if (typeof fnRefresh === 'function') { fnRefresh(); return true; }
                    if (typeof captchaRefresh === 'function') { captchaRefresh(); return true; }
                    if (typeof refreshCaptcha === 'function') { refreshCaptcha(); return true; }
                    var imgs = document.querySelectorAll('img[src*=""captcha"" i], img[src*=""cap_img"" i]');
                    for (var img of imgs) { img.src = img.src.split('?')[0] + '?t=' + Date.now(); return true; }
                    return false;
                }"));

            if (jsResult)
                _logger.LogInformation("CAPTCHA 새로고침 JS fallback 성공");
            else
                _logger.LogWarning("CAPTCHA 새로고침 방법을 찾지 못함");
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "CAPTCHA 새로고침 JS fallback 실패");
        }
    }

    private static async Task<(ILocator? inputLocator, IFrame? frame)> FindCaptchaInputAsync(
        IPage page, TimeSpan timeout, CancellationToken cancellationToken)
    {
        const string inputSelector = "input[placeholder*='문자'], input[name*='captcha' i], input[id*='captcha' i], input[name*='CAPTCHA'], input[placeholder*='보안문자'], input[placeholder*='자동입력']";
        var infinite = timeout == Timeout.InfiniteTimeSpan;
        var deadline = infinite ? DateTimeOffset.MaxValue : DateTimeOffset.UtcNow + timeout;

        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var mainInput = page.Locator(inputSelector);
                if (await mainInput.CountAsync() > 0)
                    return (mainInput, null);
            }
            catch (PlaywrightException) { }

            foreach (var frame in page.Frames)
            {
                if (frame == page.MainFrame) continue;
                try
                {
                    var frameInput = frame.Locator(inputSelector);
                    if (await frameInput.CountAsync() > 0)
                        return (frameInput, frame);
                }
                catch (PlaywrightException) { }
            }

            await Task.Delay(PollDelayMilliseconds, cancellationToken);
        }

        return (null, null);
    }

    private static OcrEngine GetNolOcrEngine()
    {
        if (_nolOcrEngine is not null)
        {
            return _nolOcrEngine;
        }

        var preferredLanguage = new Language("ko-KR");
        if (OcrEngine.IsLanguageSupported(preferredLanguage))
        {
            _nolOcrEngine = OcrEngine.TryCreateFromLanguage(preferredLanguage);
        }

        _nolOcrEngine ??= OcrEngine.TryCreateFromUserProfileLanguages();
        if (_nolOcrEngine is null)
        {
            var fallbackLanguage = OcrEngine.AvailableRecognizerLanguages.FirstOrDefault();
            if (fallbackLanguage is not null)
            {
                _nolOcrEngine = OcrEngine.TryCreateFromLanguage(fallbackLanguage);
            }
        }

        return _nolOcrEngine ?? throw new InvalidOperationException("Windows OCR 엔진을 초기화하지 못했습니다. OCR 언어 팩 설치를 확인하세요.");
    }

    private static string NormalizeNolRoundOcrText(string value)
    {
        var normalized = NormalizeText(value)
            .Replace("회차", "회", StringComparison.Ordinal)
            .Replace('희', '회')
            .Replace('히', '회')
            .Replace('외', '회');

        if (TryParseNolRound(normalized, out var round, out var time))
        {
            return $"{round} {time}";
        }

        return normalized;
    }

    private static bool TryClickNolRoundByOrder(Mat source, int offsetX, int offsetY, OpenCvSharp.Rect roundArea, string desiredRound, double scale, out string selectedRound)
    {
        selectedRound = string.Empty;
        if (!TryExtractLeadingRoundNumber(desiredRound, out var roundNumber) ||
            !int.TryParse(roundNumber, NumberStyles.None, CultureInfo.InvariantCulture, out var roundIndex) ||
            roundIndex < 1)
        {
            return false;
        }

        var candidates = FindNolRoundButtonCandidates(source, roundArea, scale);
        if (roundIndex > candidates.Count)
        {
            return false;
        }

        var target = candidates[roundIndex - 1];
        ClickAt(offsetX + target.CenterX, offsetY + target.CenterY);
        selectedRound = $"{roundIndex}회";
        return true;
    }

    private static bool TryClickNolRoundRowByIndex(int offsetX, int offsetY, OpenCvSharp.Rect roundArea, string desiredRound, double scale, int maxRows, int maxWidth, int maxHeight, out string selectedRound)
    {
        selectedRound = string.Empty;
        if (!TryExtractLeadingRoundNumber(desiredRound, out var roundNumber) ||
            !int.TryParse(roundNumber, NumberStyles.None, CultureInfo.InvariantCulture, out var roundIndex) ||
            roundIndex < 1 ||
            roundIndex > maxRows ||
            roundIndex != 1)
        {
            return false;
        }

        var rowRect = CreateNolRect(
            roundArea.Left,
            roundArea.Top + (int)Math.Round((roundIndex - 1) * NolRoundRowHeight * scale),
            roundArea.Width,
            (int)Math.Round(NolRoundRowHeight * scale),
            maxWidth,
            maxHeight);

        if (rowRect.Width <= 0 || rowRect.Height <= 0)
        {
            return false;
        }

        ClickAt(offsetX + rowRect.Left + (rowRect.Width / 2), offsetY + rowRect.Top + (rowRect.Height / 2));
        selectedRound = $"{roundIndex}회";
        return true;
    }

    private static IReadOnlyList<TemplateBounds> FindNolRoundButtonCandidates(Mat source, OpenCvSharp.Rect roundArea, double scale)
    {
        if (roundArea.Width <= 0 || roundArea.Height <= 0)
        {
            return [];
        }

        using var roi = new Mat(source, roundArea);
        using var gray = ToGray(roi);
        using var blurred = new Mat();
        Cv2.GaussianBlur(gray, blurred, new OpenCvSharp.Size(5, 5), 0);
        using var edges = new Mat();
        Cv2.Canny(blurred, edges, 50, 150);
        using var kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(3, 3));
        using var dilated = new Mat();
        Cv2.Dilate(edges, dilated, kernel, iterations: 2);

        Cv2.FindContours(dilated, out var contours, out _, RetrievalModes.External, ContourApproximationModes.ApproxSimple);

        var minWidth = Math.Max(40, (int)Math.Round(roundArea.Width * 0.28));
        var minHeight = Math.Max(20, (int)Math.Round(NolRoundRowHeight * scale * 0.45));
        var maxHeight = Math.Max(minHeight, (int)Math.Round(NolRoundRowHeight * scale * 1.8));
        var candidates = new List<TemplateBounds>();

        foreach (var contour in contours)
        {
            var rect = Cv2.BoundingRect(contour);
            if (rect.Width < minWidth || rect.Height < minHeight || rect.Height > maxHeight)
            {
                continue;
            }

            if (rect.Width > roundArea.Width || rect.Height > roundArea.Height)
            {
                continue;
            }

            var aspectRatio = rect.Width / (double)Math.Max(1, rect.Height);
            if (aspectRatio < 1.4)
            {
                continue;
            }

            candidates.Add(new TemplateBounds(roundArea.Left + rect.Left, roundArea.Top + rect.Top, rect.Width, rect.Height, 1));
        }

        if (candidates.Count == 0)
        {
            return [];
        }

        candidates.Sort((left, right) =>
        {
            var topComparison = left.Top.CompareTo(right.Top);
            return topComparison != 0 ? topComparison : left.Left.CompareTo(right.Left);
        });

        var deduplicated = new List<TemplateBounds>();
        foreach (var candidate in candidates)
        {
            if (deduplicated.Any(existing => Math.Abs(existing.CenterX - candidate.CenterX) <= Math.Max(8, candidate.Width / 5) &&
                                            Math.Abs(existing.CenterY - candidate.CenterY) <= Math.Max(8, candidate.Height / 2)))
            {
                continue;
            }

            deduplicated.Add(candidate);
        }

        return deduplicated;
    }

    private static int GetMonthDifference(DateOnly currentMonth, DateOnly targetMonth)
    {
        return ((targetMonth.Year - currentMonth.Year) * 12) + targetMonth.Month - currentMonth.Month;
    }

    private static System.Drawing.Point GetNolCalendarCellCenter(OpenCvSharp.Rect calendarRect, DateOnly desiredDate, double scale)
    {
        var firstDay = new DateOnly(desiredDate.Year, desiredDate.Month, 1);
        var startColumn = (int)firstDay.DayOfWeek;
        var index = startColumn + desiredDate.Day - 1;
        var row = index / 7;
        var column = index % 7;
        var x = calendarRect.Left + (int)Math.Round((NolCalendarGridLeftOffset + (column * NolCalendarColumnStep) + (NolCalendarCellSize / 2.0)) * scale);
        var y = calendarRect.Top + (int)Math.Round((NolCalendarGridTopOffset + (row * NolCalendarRowStep) + (NolCalendarCellSize / 2.0)) * scale);
        return new System.Drawing.Point(x, y);
    }

    private static double GetNolTemplateScale(TemplateBounds bounds, double baseWidth)
    {
        return bounds.Width / baseWidth;
    }

    private static OpenCvSharp.Rect CreateNolRect(int left, int top, int width, int height, int maxWidth, int maxHeight)
    {
        var clampedLeft = Math.Clamp(left, 0, Math.Max(0, maxWidth - 1));
        var clampedTop = Math.Clamp(top, 0, Math.Max(0, maxHeight - 1));
        var clampedWidth = Math.Clamp(width, 0, maxWidth - clampedLeft);
        var clampedHeight = Math.Clamp(height, 0, maxHeight - clampedTop);
        return new OpenCvSharp.Rect(clampedLeft, clampedTop, clampedWidth, clampedHeight);
    }

    private static async Task ClickNolScreenTemplateAsync(string fileName, double threshold, TimeSpan timeout, CancellationToken cancellationToken, string description)
    {
        var clicked = await TryClickNolScreenTemplateAsync(fileName, threshold, timeout, cancellationToken);
        if (!clicked)
        {
            throw new TimeoutException($"NOL 화면에서 {description} 템플릿을 찾지 못했습니다: {fileName}");
        }
    }

    private static async Task<bool> TryClickNolScreenTemplateAsync(string fileName, double threshold, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var template = LoadNolScreenTemplate(fileName);
        try
        {
            var deadline = DateTimeOffset.UtcNow + timeout;
            while (DateTimeOffset.UtcNow < deadline)
            {
                cancellationToken.ThrowIfCancellationRequested();

                using var frame = CaptureScreen();
                using var grayFrame = ToGray(frame.Image);
                var match = TryFindTemplateBounds(grayFrame, template, threshold, out _);
                if (match is not null)
                {
                    ClickAt(frame.OffsetX + match.Value.CenterX, frame.OffsetY + match.Value.CenterY);
                    return true;
                }

                await Task.Delay(PollDelayMilliseconds, cancellationToken);
            }

            return false;
        }
        finally
        {
            DisposeStepTemplate(template);
        }
    }

    private static StepTemplate LoadNolScreenTemplate(string fileName)
    {
        var assembly = typeof(PlaywrightTicketingAutomationService).Assembly;
        var resourceName = NolTemplateResourcePrefix + fileName;
        using var stream = assembly.GetManifestResourceStream(resourceName)
            ?? throw new FileNotFoundException($"NOL 템플릿 리소스를 찾지 못했습니다: {resourceName}");
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        using var encoded = Cv2.ImDecode(memory.ToArray(), ImreadModes.Grayscale);
        if (encoded.Empty())
        {
            throw new InvalidOperationException($"NOL 템플릿 이미지를 읽지 못했습니다: {resourceName}");
        }

        return new StepTemplate("single", null, false, BuildScaledTemplates(encoded));
    }

    private static void DisposeStepTemplate(StepTemplate template)
    {
        foreach (var scaledTemplate in template.ScaledTemplates)
        {
            scaledTemplate.Dispose();
        }
    }

    private static async Task EnsureNolPopupClosedAsync(IPage page, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var deadline = DateTimeOffset.UtcNow + timeout;

        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var closeButtons = page.Locator(".popup.is-visible .popupCloseBtn");
            var buttonCount = await closeButtons.CountAsync();
            if (buttonCount == 0)
            {
                return;
            }

            var clicked = false;
            for (var index = 0; index < buttonCount; index++)
            {
                var closeButton = closeButtons.Nth(index);
                if (!await closeButton.IsVisibleAsync())
                {
                    continue;
                }

                try
                {
                    await closeButton.ClickAsync(new LocatorClickOptions
                    {
                        Force = true,
                        Timeout = 1500
                    });
                }
                catch (PlaywrightException)
                {
                    await closeButton.EvaluateAsync("button => { button.click(); }");
                }

                clicked = true;
                await page.WaitForTimeoutAsync(150);
                break;
            }

            if (!clicked)
            {
                return;
            }

            var popupClosed = await TryDismissVisibleNolPopupsAsync(page);
            if (popupClosed)
            {
                await page.WaitForTimeoutAsync(150);
            }

            if (await page.Locator(".popup.is-visible").CountAsync() == 0)
            {
                return;
            }
        }

        throw new TimeoutException("NOL 안내 팝업을 닫지 못했습니다.");
    }

    private static async Task SelectNolDateAsync(IPage page, DateOnly desiredDate, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var side = page.Locator("#productSide");
        var calendar = side.Locator(".sideCalendar");
        var desiredMonth = new DateOnly(desiredDate.Year, desiredDate.Month, 1);

        await WaitForConditionAsync(async () => await calendar.CountAsync() > 0, timeout, cancellationToken, "NOL 달력 영역을 찾지 못했습니다.");

        for (var attempt = 0; attempt < 24; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var currentMonthText = NormalizeText(await calendar.Locator("li[data-view='month current']").InnerTextAsync());
            if (!TryParseMonth(currentMonthText, out var currentMonth))
            {
                throw new InvalidOperationException($"NOL 달력 월 텍스트를 해석하지 못했습니다: {currentMonthText}");
            }

            if (currentMonth == desiredMonth)
            {
                break;
            }

            var direction = currentMonth < desiredMonth ? "next" : "prev";
            var nav = calendar.Locator($"li[data-view='month {direction}']:not(.disabled)");
            if (await nav.CountAsync() == 0)
            {
                throw new InvalidOperationException($"NOL 달력에서 {desiredDate:yyyy.MM} 월로 이동할 수 없습니다.");
            }

            await ClickNolElementAsync(nav.First);
            await WaitForConditionAsync(
                async () => NormalizeText(await calendar.Locator("li[data-view='month current']").InnerTextAsync()) != currentMonthText,
                timeout,
                cancellationToken,
                "NOL 달력 월 전환을 확인하지 못했습니다.");
        }

        var days = calendar.Locator("ul[data-view='days'] > li");
        var count = await days.CountAsync();
        for (var index = 0; index < count; index++)
        {
            var cell = days.Nth(index);
            var dayText = NormalizeText(await cell.InnerTextAsync());
            var className = await cell.GetAttributeAsync("class") ?? string.Empty;
            if (dayText != desiredDate.Day.ToString(CultureInfo.InvariantCulture) || IsNolDisabledClass(className))
            {
                continue;
            }

            await ClickNolElementAsync(cell);
            if (!await TryWaitForConditionAsync(
                    async () => NormalizeText(await side.Locator(".containerTop .selectedData .date").InnerTextAsync()).Contains(desiredDate.ToString("yyyy.MM.dd", CultureInfo.InvariantCulture), StringComparison.Ordinal),
                    TimeSpan.FromMilliseconds(800),
                    cancellationToken))
            {
                await cell.EvaluateAsync("element => { element.scrollIntoView({ block: 'center', inline: 'center' }); element.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window })); element.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window })); element.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })); if (typeof element.click === 'function') { element.click(); } }");
            }

            await WaitForConditionAsync(
                async () => NormalizeText(await side.Locator(".containerTop .selectedData .date").InnerTextAsync()).Contains(desiredDate.ToString("yyyy.MM.dd", CultureInfo.InvariantCulture), StringComparison.Ordinal),
                timeout,
                cancellationToken,
                "NOL 관람일 선택 반영을 확인하지 못했습니다.");
            return;
        }

        throw new InvalidOperationException($"NOL 달력에서 선택 가능한 관람일을 찾지 못했습니다: {desiredDate:yyyy.MM.dd}");
    }

    private static async Task SelectNolRoundAsync(IPage page, string desiredRound, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var side = page.Locator("#productSide");
        var roundsLocator = side.Locator(".sideTimeTable .timeTableLabel[role='button']");

        await WaitForConditionAsync(async () => await roundsLocator.CountAsync() > 0, timeout, cancellationToken, "NOL 회차 목록을 찾지 못했습니다.");

        var roundItems = await page.EvaluateAsync<NolRoundItem[]>("""
            () => [...document.querySelectorAll('#productSide .sideTimeTable .timeTableLabel[role="button"]')]
                  .map((el, i) => ({ text: el.innerText, className: el.className, index: i }))
            """);

        foreach (var item in roundItems ?? [])
        {
            var roundText = NormalizeText(item.Text);
            if (!IsMatchingNolRound(roundText, desiredRound) || IsNolDisabledClass(item.ClassName))
            {
                continue;
            }

            if (item.ClassName.Contains("is-toggled", StringComparison.OrdinalIgnoreCase) &&
                await TryWaitForConditionAsync(
                    async () => IsMatchingNolRound(await side.Locator(".containerMiddle .selectedData .time").InnerTextAsync(), desiredRound),
                    TimeSpan.FromMilliseconds(300),
                    cancellationToken))
            {
                return;
            }

            await ClickNolElementAsync(roundsLocator.Nth(item.Index));
            await WaitForConditionAsync(
                async () => IsMatchingNolRound(await side.Locator(".containerMiddle .selectedData .time").InnerTextAsync(), desiredRound),
                timeout,
                cancellationToken,
                "NOL 회차 선택 반영을 확인하지 못했습니다.");
            return;
        }

        throw new InvalidOperationException($"NOL 회차를 찾지 못했습니다: {desiredRound}");
    }

    private sealed class NolRoundItem
    {
        public string Text { get; set; } = string.Empty;
        public string ClassName { get; set; } = string.Empty;
        public int Index { get; set; }
    }

    private static async Task<IPage> ClickNolBookingAsync(IPage page, TimeSpan timeout, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var bookingButton = page.Locator("#productSide a.sideBtn.is-primary").First;
        if (await bookingButton.CountAsync() == 0)
        {
            throw new InvalidOperationException("NOL 예매하기 버튼을 찾지 못했습니다.");
        }

        var beforeUrl = page.Url;
        var beforeTitle = await GetPageTitleOrEmptyAsync(page);
        var beforePages = page.Context.Pages.Where(x => !x.IsClosed).ToHashSet();
        await bookingButton.ScrollIntoViewIfNeededAsync();
        await ClickNolBookingButtonAsync(bookingButton, timeout);

        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var openPages = page.Context.Pages.Where(x => !x.IsClosed).ToList();
            var newPage = openPages.FirstOrDefault(x => !beforePages.Contains(x));
            if (newPage is not null)
            {
                return await PrepareNolBookingResultPageAsync(newPage, deadline);
            }

            if (!page.IsClosed && await HasNolBookingResultAppearedAsync(page, beforeUrl, beforeTitle))
            {
                return await PrepareNolBookingResultPageAsync(page, deadline);
            }

            if (page.IsClosed && openPages.Count > 0)
            {
                var fallbackPage = openPages[^1];
                return await PrepareNolBookingResultPageAsync(fallbackPage, deadline);
            }

            await Task.Delay(PollDelayMilliseconds, cancellationToken);
        }

        throw new TimeoutException("NOL 예매하기 클릭 후 페이지 전환을 확인하지 못했습니다.");
    }

    private static async Task ClickNolBookingButtonAsync(ILocator bookingButton, TimeSpan timeout)
    {
        var clickTimeout = (float)Math.Min(timeout.TotalMilliseconds, 5000);
        PlaywrightException? firstClickException = null;

        try
        {
            await bookingButton.ClickAsync(new LocatorClickOptions
            {
                Timeout = clickTimeout
            });
            return;
        }
        catch (PlaywrightException ex)
        {
            firstClickException = ex;
        }

        try
        {
            await bookingButton.ClickAsync(new LocatorClickOptions
            {
                Force = true,
                Timeout = clickTimeout
            });
        }
        catch (PlaywrightException ex) when (firstClickException is not null)
        {
            throw new InvalidOperationException(
                $"NOL 예매하기 버튼 클릭에 실패했습니다. firstAttempt={firstClickException.Message}; secondAttempt={ex.Message}",
                ex);
        }
    }

    private static async Task<IPage> PrepareNolBookingResultPageAsync(IPage page, DateTimeOffset deadline)
    {
        await page.WaitForLoadStateAsync(LoadState.DOMContentLoaded, new PageWaitForLoadStateOptions
        {
            Timeout = (float)Math.Max(1000, (deadline - DateTimeOffset.UtcNow).TotalMilliseconds)
        });
        await page.BringToFrontAsync();
        return page;
    }

    private static async Task<bool> HasNolBookingResultAppearedAsync(IPage page, string beforeUrl, string beforeTitle)
    {
        if (!string.Equals(page.Url, beforeUrl, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var currentTitle = await GetPageTitleOrEmptyAsync(page);
        return !string.IsNullOrWhiteSpace(currentTitle) &&
               !string.Equals(currentTitle, beforeTitle, StringComparison.Ordinal) &&
               await page.Locator("#productSide").CountAsync() == 0;
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

    private static bool IsNolDisabledClass(string className)
    {
        return className.Contains("disabled", StringComparison.OrdinalIgnoreCase) ||
               className.Contains("muted", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsMatchingNolRound(string actual, string desired)
    {
        var normalizedActual = NormalizeNolRoundOcrText(actual);
        var normalizedDesired = NormalizeNolRoundOcrText(desired);

        if (TryParseNolRound(normalizedActual, out var actualRound, out var actualTime) &&
            TryParseNolRound(normalizedDesired, out var desiredRound, out var desiredTime))
        {
            return string.Equals(actualRound, desiredRound, StringComparison.Ordinal) &&
                   string.Equals(actualTime, desiredTime, StringComparison.Ordinal);
        }

        if (TryParseNolRound(normalizedActual, out actualRound, out actualTime) && TryParseNolRoundLabel(normalizedDesired, out desiredRound))
        {
            return string.Equals(actualRound, desiredRound, StringComparison.Ordinal);
        }

        if (TryParseNolRoundLabel(normalizedActual, out actualRound) && TryParseNolRoundLabel(normalizedDesired, out desiredRound))
        {
            return string.Equals(actualRound, desiredRound, StringComparison.Ordinal);
        }

        if (TryExtractLeadingRoundNumber(normalizedActual, out var actualRoundNumber) &&
            TryExtractLeadingRoundNumber(normalizedDesired, out var desiredRoundNumber))
        {
            return string.Equals(actualRoundNumber, desiredRoundNumber, StringComparison.Ordinal);
        }

        return string.Equals(NormalizeText(normalizedActual), NormalizeText(normalizedDesired), StringComparison.Ordinal);
    }

    private static bool TryParseNolRound(string value, out string round, out string time)
    {
        round = string.Empty;
        time = string.Empty;
        var normalized = NormalizeText(value)
            .Replace("회차", "회", StringComparison.Ordinal)
            .Replace('희', '회')
            .Replace('히', '회')
            .Replace('외', '회');
        var match = NolRoundPattern.Match(normalized);
        if (!match.Success)
        {
            return false;
        }

        round = $"{NormalizeText(match.Groups["round"].Value)}회";
        time = NormalizeNolTime(match.Groups["time"].Value);
        return !string.IsNullOrWhiteSpace(round) && !string.IsNullOrWhiteSpace(time);
    }

    private static bool TryParseNolRoundLabel(string value, out string round)
    {
        round = string.Empty;
        var normalized = NormalizeText(value).Replace("회차", "회", StringComparison.Ordinal);
        if (Regex.IsMatch(normalized, "^\\d+$", RegexOptions.Compiled))
        {
            round = $"{normalized}회";
            return true;
        }

        var match = Regex.Match(normalized, @"(?<round>\d+\s*회)", RegexOptions.Compiled);
        if (!match.Success)
        {
            return false;
        }

        round = NormalizeText(match.Groups["round"].Value);
        return !string.IsNullOrWhiteSpace(round);
    }

    private static bool TryExtractLeadingRoundNumber(string value, out string roundNumber)
    {
        roundNumber = string.Empty;
        var match = Regex.Match(NormalizeText(value), @"^\D*(?<round>\d{1,2})(?:\D|$)", RegexOptions.Compiled);
        if (!match.Success)
        {
            return false;
        }

        roundNumber = NormalizeText(match.Groups["round"].Value);
        return !string.IsNullOrWhiteSpace(roundNumber);
    }

    private static string NormalizeNolTime(string value)
    {
        var digits = DigitsOnlyPattern.Replace(value, string.Empty);
        if (digits.Length == 3)
        {
            return $"{digits[0]}:{digits[1..]}";
        }

        if (digits.Length == 4)
        {
            return $"{digits[..2]}:{digits[2..]}";
        }

        return NormalizeText(value).Replace('.', ':').Replace(',', ':');
    }

    private static async Task<bool> TryDismissVisibleNolPopupsAsync(IPage page)
    {
        var dismissed = await page.EvaluateAsync<bool>("""
            () => {
                const buttons = Array.from(document.querySelectorAll('.popup.is-visible .popupCloseBtn'));
                for (const button of buttons) {
                    button.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
                    if (typeof button.click === 'function') {
                        button.click();
                    }
                }

                return buttons.length > 0;
            }
            """);

        return dismissed;
    }

    private static bool IsClosedTargetError(PlaywrightException ex)
    {
        return ex.Message.Contains("Target page, context or browser has been closed", StringComparison.OrdinalIgnoreCase) ||
               ex.Message.Contains("Browser has been closed", StringComparison.OrdinalIgnoreCase) ||
               ex.Message.Contains("Target closed", StringComparison.OrdinalIgnoreCase);
    }

    private static async Task ClickNolElementAsync(ILocator locator)
    {
        try
        {
            await locator.ClickAsync(new LocatorClickOptions
            {
                Force = true,
                Timeout = 1500
            });
        }
        catch (PlaywrightException)
        {
            await locator.EvaluateAsync("element => { element.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })); if (typeof element.click === 'function') { element.click(); } }");
        }
    }



    private static async Task<string> DescribeNolPageStateAsync(IPage? page)
    {
        if (page is null)
        {
            return "page=null";
        }

        try
        {
            var popupCount = await page.Locator(".popup.is-visible").CountAsync();
            var currentUrl = page.Url;
            var currentTitle = await GetPageTitleOrEmptyAsync(page);
            var pageCount = page.Context.Pages.Count;
            var productSideCount = await page.Locator("#productSide").CountAsync();
            var roundButtonCount = await page.Locator(".sideTimeTable .timeTableLabel[role='button']").CountAsync();
            var bookingButtonCount = await page.Locator("#productSide a.sideBtn.is-primary").CountAsync();
            var selectedDateText = await GetLocatorTextOrEmptyAsync(page.Locator("#productSide .containerTop .selectedData .date").First);
            var selectedRoundText = await GetLocatorTextOrEmptyAsync(page.Locator("#productSide .containerMiddle .selectedData .time").First);
            var openPages = string.Join(", ",
                page.Context.Pages.Select((x, index) => $"[{index}]closed={x.IsClosed};url={SafePageUrl(x)}"));
            return $"pageClosed={page.IsClosed}, url={currentUrl}, title={currentTitle}, popupCount={popupCount}, productSideCount={productSideCount}, roundButtonCount={roundButtonCount}, bookingButtonCount={bookingButtonCount}, selectedDate={selectedDateText}, selectedRound={selectedRoundText}, contextPageCount={pageCount}, openPages={openPages}";
        }
        catch (PlaywrightException ex) when (IsClosedTargetError(ex))
        {
            return $"page-state-unavailable:{ex.Message}";
        }
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
