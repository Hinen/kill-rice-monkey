using Microsoft.Extensions.Logging;
using Microsoft.Playwright;
using OpenCvSharp;
using Polly;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;
using KillRiceMonkey.Application.Abstractions;
using KillRiceMonkey.Application.Models;
using StepTemplate = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.StepTemplate;
using TemplateBounds = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.TemplateBounds;

namespace KillRiceMonkey.Infrastructure.Services;

public sealed class NolAutomationService : INolAutomationService, IAsyncDisposable
{
    private static readonly Regex NolRoundPattern = new(@"^\D*(?<round>\d{1,2})\s*(?:회차|회|희|히|외)?\s*(?<time>\d{1,2}(?::|\.|,)?\d{2})", RegexOptions.Compiled);
    private const string NolRemoteDebugLaunchUrl = "https://tickets.interpark.com/";
    private const string NolCdpEndpoint = "http://127.0.0.1:9222/";
    private const string NolTemplateResourcePrefix = PlaywrightRuntime.EmbeddedTemplatePrefix + "Nol.";
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
    private readonly ILogger<NolAutomationService> _logger;
    private readonly PlaywrightRuntime _runtime;
    private readonly ResiliencePipeline _pipeline;
    private readonly SemaphoreSlim _nolBrowserLock = new(1, 1);
    private IBrowser? _preparedNolConnectedBrowser;
    private IPage? _preparedNolPage;
    private bool _popupClosedDuringPrepare;

    public NolAutomationService(
        ILogger<NolAutomationService> logger,
        PlaywrightRuntime runtime)
    {
        _logger = logger;
        _runtime = runtime;
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
                return await RunNolAutomationAsync(request, progress, token);
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
        _nolBrowserLock.Dispose();
    }

    public async Task<bool> IsRemoteDebugBrowserAvailableAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return await PlaywrightRuntime.IsCdpEndpointAvailableAsync(NolCdpEndpoint, cancellationToken);
    }

    public async Task<bool> IsAutomationPreparedAsync(CancellationToken cancellationToken)
    {
        return await TryGetPreparedNolPageAsync(cancellationToken) is not null;
    }

    public async Task<string> LaunchRemoteDebugBrowserAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (await PlaywrightRuntime.IsCdpEndpointAvailableAsync(NolCdpEndpoint, cancellationToken))
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
            if (await PlaywrightRuntime.IsCdpEndpointAvailableAsync(NolCdpEndpoint, cancellationToken))
            {
                return $"remote debug 브라우저를 열었습니다: {NolRemoteDebugLaunchUrl}";
            }

            await Task.Delay(200, cancellationToken);
        }

        return $"브라우저 실행은 요청했지만 remote debug 포트 확인이 지연되고 있습니다. 직접 확인: {NolCdpEndpoint}";
    }

    public async Task<string> PrepareAutomationAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!await IsRemoteDebugBrowserAvailableAsync(cancellationToken))
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
        return $"NOL 준비 완료: {PlaywrightRuntime.SafePageUrl(page)}";
    }


    private async Task<AutomationRunResult> RunNolAutomationAsync(TicketingJobRequest request, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        if (request.MatchThreshold <= 0 || request.MatchThreshold > 1)
        {
            return new AutomationRunResult(false, "매칭 임계값은 0보다 크고 1 이하여야 합니다.", DateTimeOffset.Now);
        }

        if (request.StepTimeoutSeconds <= 0)
        {
            return new AutomationRunResult(false, "단계별 제한 시간은 1초 이상이어야 합니다.", DateTimeOffset.Now);
        }

        if (!PlaywrightRuntime.TryParseDesiredDate(request.DesiredDate, out var desiredDate))
        {
            return new AutomationRunResult(false, "관람일 형식이 올바르지 않습니다. 예: 2026.04.11", DateTimeOffset.Now);
        }

        if (string.IsNullOrWhiteSpace(request.DesiredRound))
        {
            return new AutomationRunResult(false, "회차 값이 비어 있습니다. 예: 1회 19:00", DateTimeOffset.Now);
        }

        var desiredRound = PlaywrightRuntime.NormalizeText(request.DesiredRound);
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
                PlaywrightRuntime.GetLogDirectoryPath());
        }

        try
        {
            _logger.LogInformation("NOL screen automation started. runId={RunId}, date={Date}, round={Round}", runId, desiredDate, desiredRound);

            SetStage("attach-existing-browser");
            var cdpResult = await TryRunNolAutomationViaConnectedBrowserAsync(desiredDate, desiredRound, timeout, progress, cancellationToken);
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
            _logger.LogError(ex, "NOL screen automation failed. runId={RunId}, phase={Phase}, date={Date}, round={Round}, logDir={LogDir}", runId, phase, desiredDate, desiredRound, PlaywrightRuntime.GetLogDirectoryPath());
            return new AutomationRunResult(false, $"NOL 화면 자동화 시간 초과: {ex.Message} | 상세 로그: {PlaywrightRuntime.GetLogDirectoryPath()}", DateTimeOffset.Now);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "NOL screen automation failed. runId={RunId}, phase={Phase}, date={Date}, round={Round}, logDir={LogDir}", runId, phase, desiredDate, desiredRound, PlaywrightRuntime.GetLogDirectoryPath());
            return new AutomationRunResult(false, $"NOL 화면 자동화 예외: {ex.Message} | 상세 로그: {PlaywrightRuntime.GetLogDirectoryPath()}", DateTimeOffset.Now);
        }
    }

    private async Task<AutomationRunResult?> TryRunNolAutomationViaConnectedBrowserAsync(DateOnly desiredDate, string desiredRound, TimeSpan timeout, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
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

            _logger.LogInformation("Connected-browser NOL automation selected page. url={Url}", PlaywrightRuntime.SafePageUrl(page));
            await page.BringToFrontAsync();
            if (!_popupClosedDuringPrepare)
            {
                await EnsureNolPopupClosedAsync(page, TimeSpan.FromSeconds(2), cancellationToken);
            }
            _popupClosedDuringPrepare = false;
            progress?.Report(new AutomationProgress("날짜 선택 중"));
            await SelectNolDateAsync(page, desiredDate, timeout, cancellationToken);
            progress?.Report(new AutomationProgress("날짜 선택 완료", "날짜 선택 완료"));
            progress?.Report(new AutomationProgress("회차 선택 중"));
            await SelectNolRoundAsync(page, desiredRound, timeout, cancellationToken);
            progress?.Report(new AutomationProgress("회차 선택 완료", "회차 선택 완료"));
            progress?.Report(new AutomationProgress("예매 클릭 중"));
            var captchaPage = await ClickNolBookingAsync(page, timeout, progress, cancellationToken);
            progress?.Report(new AutomationProgress("예매 클릭 완료", "예매 클릭 → 대기열/캡차 진입"));
            progress?.Report(new AutomationProgress("캡차 입력 중"));
            await SolveNolCaptchaAsync(captchaPage, timeout, cancellationToken);
            progress?.Report(new AutomationProgress("캡차 입력 완료", "캡차 처리 완료"));
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

            _preparedNolConnectedBrowser = await _runtime.TryConnectToExistingChromiumBrowserAsync(NolCdpEndpoint, cancellationToken);
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
            return PlaywrightRuntime.SafePageUrl(page).Contains("tickets.interpark.com/goods/", StringComparison.OrdinalIgnoreCase) &&
                   await page.Locator("#productSide").CountAsync() > 0;
        }
        catch (PlaywrightException ex) when (PlaywrightRuntime.IsClosedTargetError(ex))
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

    public async Task<bool> IsPageReadyAsync(CancellationToken cancellationToken)
    {
        if (_preparedNolPage is not null && !_preparedNolPage.IsClosed)
        {
            try
            {
                if (PlaywrightRuntime.SafePageUrl(_preparedNolPage).Contains("tickets.interpark.com/goods/", StringComparison.OrdinalIgnoreCase) &&
                    await _preparedNolPage.Locator("#productSide").CountAsync() > 0)
                {
                    return true;
                }
            }
            catch (PlaywrightException)
            {
            }
        }

        if (_preparedNolConnectedBrowser is null)
        {
            try
            {
                _preparedNolConnectedBrowser = await _runtime.TryConnectToExistingChromiumBrowserAsync(NolCdpEndpoint, cancellationToken);
            }
            catch
            {
                return false;
            }

            if (_preparedNolConnectedBrowser is null)
            {
                return false;
            }
        }

        foreach (var page in _preparedNolConnectedBrowser.Contexts.SelectMany(x => x.Pages).Where(x => !x.IsClosed))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var url = PlaywrightRuntime.SafePageUrl(page);
            if (!url.Contains("tickets.interpark.com/goods/", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            try
            {
                if (await page.Locator("#productSide").CountAsync() > 0)
                {
                    _preparedNolPage = page;
                    return true;
                }
            }
            catch (PlaywrightException)
            {
            }
        }

        return false;
    }

    private static async Task<IPage?> FindConnectedNolPageAsync(IBrowser browser, CancellationToken cancellationToken)
    {
        var candidates = new List<(IPage Page, int Score)>();
        foreach (var page in browser.Contexts.SelectMany(x => x.Pages).Where(x => !x.IsClosed))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var url = PlaywrightRuntime.SafePageUrl(page);
            var score = 0;
            if (url.Contains("tickets.interpark.com", StringComparison.OrdinalIgnoreCase))
            {
                score += 2;
            }

            if (url.Contains("/goods/", StringComparison.OrdinalIgnoreCase))
            {
                score += 4;
            }

            if (await PlaywrightRuntime.TryWaitForConditionAsync(async () => await page.Locator("#productSide").CountAsync() > 0, TimeSpan.FromMilliseconds(400), cancellationToken))
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
            catch (PlaywrightException ex) when (PlaywrightRuntime.IsClosedTargetError(ex))
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

                using var frame = PlaywrightRuntime.CaptureScreen();
                using var grayFrame = PlaywrightRuntime.ToGray(frame.Image);
                var anchor = PlaywrightRuntime.TryFindTemplateBounds(grayFrame, template, threshold, out _);
                if (anchor is null)
                {
                    await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
                    continue;
                }

                var scale = GetNolTemplateScale(anchor.Value, NolPanelToggleWidth);
                var calendarRect = CreateNolRect(anchor.Value.Left + (int)Math.Round(NolCalendarLeftOffset * scale), anchor.Value.Top + (int)Math.Round(NolCalendarTopOffset * scale), (int)Math.Round(NolCalendarWidth * scale), (int)Math.Round(NolCalendarHeight * scale), frame.Image.Width, frame.Image.Height);
                var monthRect = CreateNolRect(calendarRect.Left, calendarRect.Top, calendarRect.Width, (int)Math.Round(NolCalendarMonthHeaderHeight * scale), frame.Image.Width, frame.Image.Height);
                var monthText = await ReadNolOcrTextAsync(frame.Image, monthRect, cancellationToken);
                if (!PlaywrightRuntime.TryParseMonth(monthText, out var displayedMonth))
                {
                    await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
                    continue;
                }

                var targetMonth = new DateOnly(desiredDate.Year, desiredDate.Month, 1);
                var monthDifference = GetMonthDifference(displayedMonth, targetMonth);
                if (monthDifference == 0)
                {
                    var cellCenter = GetNolCalendarCellCenter(calendarRect, desiredDate, scale);
                    PlaywrightRuntime.ClickAt(frame.OffsetX + cellCenter.X, frame.OffsetY + cellCenter.Y);
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
                PlaywrightRuntime.ClickAt(frame.OffsetX + arrowX, frame.OffsetY + arrowY);
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

                using var frame = PlaywrightRuntime.CaptureScreen();
                using var grayFrame = PlaywrightRuntime.ToGray(frame.Image);
                var anchor = PlaywrightRuntime.TryFindTemplateBounds(grayFrame, roundHeaderTemplate, threshold, out _);
                if (anchor is null)
                {
                    await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
                    continue;
                }

                var scale = GetNolTemplateScale(anchor.Value, NolPanelToggleWidth);
                var bookingBounds = PlaywrightRuntime.TryFindTemplateBounds(grayFrame, bookingButtonTemplate, Math.Min(0.82, threshold + 0.02), out _);
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

                    PlaywrightRuntime.ClickAt(frame.OffsetX + rowRect.Left + (rowRect.Width / 2), frame.OffsetY + rowRect.Top + (rowRect.Height / 2));
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

                await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
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
        var text = await PlaywrightRuntime.RecognizeTextAsync(roi, applyThreshold: false, cancellationToken);
        if (!string.IsNullOrWhiteSpace(text))
        {
            return text;
        }

        return await PlaywrightRuntime.RecognizeTextAsync(roi, applyThreshold: true, cancellationToken);
    }

    private static string NormalizeNolRoundOcrText(string value)
    {
        var normalized = PlaywrightRuntime.NormalizeText(value)
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
        PlaywrightRuntime.ClickAt(offsetX + target.CenterX, offsetY + target.CenterY);
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

        PlaywrightRuntime.ClickAt(offsetX + rowRect.Left + (rowRect.Width / 2), offsetY + rowRect.Top + (rowRect.Height / 2));
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
        using var gray = PlaywrightRuntime.ToGray(roi);
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

                using var frame = PlaywrightRuntime.CaptureScreen();
                using var grayFrame = PlaywrightRuntime.ToGray(frame.Image);
                var match = PlaywrightRuntime.TryFindTemplateBounds(grayFrame, template, threshold, out _);
                if (match is not null)
                {
                    PlaywrightRuntime.ClickAt(frame.OffsetX + match.Value.CenterX, frame.OffsetY + match.Value.CenterY);
                    return true;
                }

                await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
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

        return new StepTemplate("single", null, false, PlaywrightRuntime.BuildScaledTemplates(encoded));
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

        await PlaywrightRuntime.WaitForConditionAsync(async () => await calendar.CountAsync() > 0, timeout, cancellationToken, "NOL 달력 영역을 찾지 못했습니다.");

        for (var attempt = 0; attempt < 24; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var currentMonthText = PlaywrightRuntime.NormalizeText(await calendar.Locator("li[data-view='month current']").InnerTextAsync());
            if (!PlaywrightRuntime.TryParseMonth(currentMonthText, out var currentMonth))
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

            await PlaywrightRuntime.ClickElementAsync(nav.First);
            await PlaywrightRuntime.WaitForConditionAsync(
                async () => PlaywrightRuntime.NormalizeText(await calendar.Locator("li[data-view='month current']").InnerTextAsync()) != currentMonthText,
                timeout,
                cancellationToken,
                "NOL 달력 월 전환을 확인하지 못했습니다.");
        }

        var days = calendar.Locator("ul[data-view='days'] > li");
        var count = await days.CountAsync();
        for (var index = 0; index < count; index++)
        {
            var cell = days.Nth(index);
            var dayText = PlaywrightRuntime.NormalizeText(await cell.InnerTextAsync());
            var className = await cell.GetAttributeAsync("class") ?? string.Empty;
            if (dayText != desiredDate.Day.ToString(CultureInfo.InvariantCulture) || IsNolDisabledClass(className))
            {
                continue;
            }

            await PlaywrightRuntime.ClickElementAsync(cell);
            if (!await PlaywrightRuntime.TryWaitForConditionAsync(
                    async () => PlaywrightRuntime.NormalizeText(await side.Locator(".containerTop .selectedData .date").InnerTextAsync()).Contains(desiredDate.ToString("yyyy.MM.dd", CultureInfo.InvariantCulture), StringComparison.Ordinal),
                    TimeSpan.FromMilliseconds(800),
                    cancellationToken))
            {
                await cell.EvaluateAsync("element => { element.scrollIntoView({ block: 'center', inline: 'center' }); element.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window })); element.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window })); element.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })); if (typeof element.click === 'function') { element.click(); } }");
            }

            await PlaywrightRuntime.WaitForConditionAsync(
                async () => PlaywrightRuntime.NormalizeText(await side.Locator(".containerTop .selectedData .date").InnerTextAsync()).Contains(desiredDate.ToString("yyyy.MM.dd", CultureInfo.InvariantCulture), StringComparison.Ordinal),
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

        await PlaywrightRuntime.WaitForConditionAsync(async () => await roundsLocator.CountAsync() > 0, timeout, cancellationToken, "NOL 회차 목록을 찾지 못했습니다.");

        var roundItems = await page.EvaluateAsync<NolRoundItem[]>("""
            () => [...document.querySelectorAll('#productSide .sideTimeTable .timeTableLabel[role="button"]')]
                  .map((el, i) => ({ text: el.innerText, className: el.className, index: i }))
            """);

        foreach (var item in roundItems ?? [])
        {
            var roundText = PlaywrightRuntime.NormalizeText(item.Text);
            if (!IsMatchingNolRound(roundText, desiredRound) || IsNolDisabledClass(item.ClassName))
            {
                continue;
            }

            if (item.ClassName.Contains("is-toggled", StringComparison.OrdinalIgnoreCase) &&
                await PlaywrightRuntime.TryWaitForConditionAsync(
                    async () => IsMatchingNolRound(await side.Locator(".containerMiddle .selectedData .time").InnerTextAsync(), desiredRound),
                    TimeSpan.FromMilliseconds(300),
                    cancellationToken))
            {
                return;
            }

            await PlaywrightRuntime.ClickElementAsync(roundsLocator.Nth(item.Index));
            await PlaywrightRuntime.WaitForConditionAsync(
                async () => IsMatchingNolRound(await side.Locator(".containerMiddle .selectedData .time").InnerTextAsync(), desiredRound),
                timeout,
                cancellationToken,
                "NOL 회차 선택 반영을 확인하지 못했습니다.");
            return;
        }

        throw new InvalidOperationException($"NOL 회차를 찾지 못했습니다: {desiredRound}");
    }
    private async Task SolveNolCaptchaAsync(IPage page, TimeSpan timeout, CancellationToken cancellationToken)
    {
        const int maxAttempts = 5;
        const int melonCaptchaLength = 6;

        _logger.LogInformation("CAPTCHA 입력창 대기 시작 (최대 500ms). url={Url}", PlaywrightRuntime.SafePageUrl(page));
        var (inputLocator, captchaFrame) = await FindNolCaptchaInputAsync(page, TimeSpan.FromMilliseconds(500), cancellationToken);
        if (inputLocator is null)
        {
            foreach (var contextPage in page.Context.Pages.Where(p => p != page && !p.IsClosed))
            {
                (inputLocator, captchaFrame) = await FindNolCaptchaInputAsync(contextPage, TimeSpan.FromMilliseconds(300), cancellationToken);
                if (inputLocator is not null)
                {
                    page = contextPage;
                    break;
                }
            }
        }

        if (inputLocator is null)
        {
            _logger.LogInformation("CAPTCHA 입력창 없음 — CAPTCHA 없는 공연으로 판단, 좌석 선택으로 진행.");
            return;
        }

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (page.IsClosed) return;

            var attemptSw = Stopwatch.StartNew();

            string text;
            try
            {
                text = await _runtime.RecognizeCaptchaTextAsync(inputLocator, page, captchaFrame, cancellationToken);
            }
            catch (Exception ocrEx) when (ocrEx is PlaywrightException or TimeoutException)
            {
                _logger.LogWarning("[CAPTCHA] OCR 실패 (attempt={Attempt}): {Message}", attempt, ocrEx.Message);
                if (page.IsClosed) return;

                var inputGone = false;
                try { inputGone = await inputLocator.CountAsync() == 0; } catch { }
                if (inputGone) { _logger.LogInformation("[CAPTCHA] input 사라짐 — 좌석 진행."); return; }

                bool captchaHidden = false;
                try
                {
                    captchaHidden = await page.EvaluateAsync<bool>(@"() => {
                        const el = document.querySelector('.captcha_area, .wrap_captcha, #divRecaptcha, [class*=""captcha""]');
                        if (!el) return true;
                        const s = window.getComputedStyle(el);
                        return s.display === 'none' || s.visibility === 'hidden' || s.opacity === '0';
                    }");
                }
                catch { }

                if (captchaHidden)
                {
                    _logger.LogInformation("[CAPTCHA] CAPTCHA 영역 hidden 감지 — 통과로 진행. attempt={Attempt}", attempt);
                    return;
                }

                if (attempt < maxAttempts)
                    await TryRefreshNolCaptchaImageAsync(page, captchaFrame, cancellationToken);
                continue;
            }

            if (string.IsNullOrEmpty(text))
            {
                _logger.LogInformation("[CAPTCHA] OCR 빈 결과 (attempt={Attempt}, {Ms}ms) — 이미지 없음, 좌석 진행.", attempt, attemptSw.ElapsedMilliseconds);
                return;
            }

            _logger.LogInformation("CAPTCHA attempt {Attempt}/{Max}: text={Text} ocrMs={OcrMs}", attempt, maxAttempts, text, attemptSw.ElapsedMilliseconds);

            if (text.Length != melonCaptchaLength)
            {
                if (attempt < maxAttempts)
                    await TryRefreshNolCaptchaImageAsync(page, captchaFrame, cancellationToken);
                continue;
            }

            try { await page.EvaluateAsync("() => { window.__melonAlertDetected = false; }"); } catch { }

            try
            {
                await inputLocator.First.FillAsync(text, new LocatorFillOptions { Timeout = 500 });
            }
            catch (TimeoutException)
            {
                try
                {
                    await inputLocator.First.EvaluateAsync(@"(el, val) => {
                        var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
                        if (setter) { setter.call(el, val); } else { el.value = val; }
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        if (typeof jQuery !== 'undefined') { jQuery(el).val(val).trigger('input').trigger('change'); }
                    }", text);
                }
                catch (PlaywrightException ex)
                {
                    _logger.LogWarning(ex, "CAPTCHA 입력 실패");
                    continue;
                }
            }

            const string submitSelector = "button:text-is('입력완료'), a:has-text('입력완료'), a[onclick*='fnCheck']";
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

            if (submitLocator is not null && submitCount > 0)
            {
                try
                {
                    await submitLocator.First.EvaluateAsync(@"el => {
                        if (el.disabled) el.disabled = false;
                        el.click();
                    }");
                }
                catch (PlaywrightException)
                {
                    try { await submitLocator.First.ClickAsync(new LocatorClickOptions { Timeout = 500, Force = true }); }
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
                        catch (PlaywrightException) { try { await inputLocator.First.PressAsync("Enter"); } catch (PlaywrightException) { } }
                    }
                }
            }
            else
            {
                try
                {
                    await inputLocator.First.EvaluateAsync(@"el => {
                        if (typeof fnCheck === 'function') { fnCheck(); }
                        else if (el.form) { el.form.submit(); }
                        else { el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, bubbles: true })); }
                    }");
                }
                catch (PlaywrightException) { try { await inputLocator.First.PressAsync("Enter"); } catch (PlaywrightException) { } }
            }

            _logger.LogInformation("[CAPTCHA] submit 완료. alert/input 확인 시작. attempt={Attempt}", attempt);

            var alertDetected = false;
            try { alertDetected = await page.EvaluateAsync<bool>("() => window.__melonAlertDetected === true"); } catch { }

            if (alertDetected)
            {
                _logger.LogInformation("[CAPTCHA] alert 감지 — 틀린 CAPTCHA, 재시도. attempt={Attempt}", attempt);
                try { await page.EvaluateAsync("() => { window.__melonAlertDetected = false; }"); } catch { }
                if (attempt < maxAttempts)
                    await TryRefreshNolCaptchaImageAsync(page, captchaFrame, cancellationToken);
                continue;
            }

            _logger.LogInformation("[CAPTCHA] CAPTCHA 제출 완료 (alert 없음). 좌석 진행. attempt={Attempt}, totalMs={Ms}", attempt, attemptSw.ElapsedMilliseconds);
            return;
        }

        _logger.LogWarning("[CAPTCHA] CAPTCHA 자동 인식 {Max}회 모두 실패 — 좌석 선택 진행 시도.", maxAttempts);
    }

    private async Task TryRefreshNolCaptchaImageAsync(IPage page, IFrame? captchaFrame, CancellationToken cancellationToken)
    {
        if (page.IsClosed)
            return;

        ILocator FrameOrPage(string selector) =>
            captchaFrame is not null ? captchaFrame.Locator(selector) : page.Locator(selector);

        async Task<bool> EvalJs(string script)
        {
            return captchaFrame is not null
                ? await captchaFrame.EvaluateAsync<bool>(script)
                : await page.EvaluateAsync<bool>(script);
        }

        try
        {
            var jsResult = await EvalJs(@"() => {
                if (typeof fnCapchaRefresh === 'function') { fnCapchaRefresh(); return true; }
                if (typeof fnRefresh === 'function') { fnRefresh(); return true; }
                if (typeof captchaRefresh === 'function') { captchaRefresh(); return true; }
                if (typeof refreshCaptcha === 'function') { refreshCaptcha(); return true; }
                return false;
            }");

            if (jsResult)
            {
                await Task.Delay(50, cancellationToken);
                return;
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "CAPTCHA JS 새로고침 실패");
        }

        const string refreshSelector = "#divRecaptcha .capchaBtns a:last-of-type, #btnReload, .refreshBtn, [class*='buttonRefresh'], button[aria-label*='새 문자']";
        try
        {
            var refreshLocator = FrameOrPage(refreshSelector);
            var count = await refreshLocator.CountAsync();
            if (count > 0)
            {
                await refreshLocator.First.ClickAsync(new LocatorClickOptions { Timeout = 500, Force = true });
                await Task.Delay(50, cancellationToken);
                return;
            }
        }
        catch (Exception ex)
        {
            _logger.LogDebug(ex, "CAPTCHA 새로고침 버튼 클릭 실패");
        }

        try
        {
            await EvalJs(@"() => {
                var imgs = document.querySelectorAll('#imgCaptcha, #captchaImg, img[src*=""captcha"" i], img[src*=""cap_img"" i]');
                for (var img of imgs) { img.src = img.src.split('?')[0] + '?t=' + Date.now(); }
                return true;
            }");
            await Task.Delay(50, cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "CAPTCHA 새로고침 모든 방법 실패");
        }
    }

    private static async Task<(ILocator? inputLocator, IFrame? frame)> FindNolCaptchaInputAsync(
        IPage page, TimeSpan timeout, CancellationToken cancellationToken)
    {
        const string inputSelector = "#txtCaptcha, [class*='captchaInput'] input, [class*='captchaInput'], input[placeholder*='문자'], input[name*='captcha' i], input[id*='captcha' i], input[name*='CAPTCHA'], input[placeholder*='보안문자'], input[placeholder*='자동입력']";
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

            await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
        }

        return (null, null);
    }

    private sealed class NolRoundItem
    {
        public string Text { get; set; } = string.Empty;
        public string ClassName { get; set; } = string.Empty;
        public int Index { get; set; }
    }

    private static async Task<IPage> ClickNolBookingAsync(IPage page, TimeSpan timeout, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var bookingButton = page.Locator("#productSide a.sideBtn.is-primary").First;
        if (await bookingButton.CountAsync() == 0)
        {
            throw new InvalidOperationException("NOL 예매하기 버튼을 찾지 못했습니다.");
        }

        var beforeUrl = page.Url;
        var beforeTitle = await PlaywrightRuntime.GetPageTitleOrEmptyAsync(page);
        var beforePages = page.Context.Pages.Where(x => !x.IsClosed).ToHashSet();
        await bookingButton.ScrollIntoViewIfNeededAsync();
        await PlaywrightRuntime.ClickBookingButtonAsync(bookingButton, timeout);

        var deadline = DateTimeOffset.UtcNow + timeout;
        var pageTransitionDetected = false;

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

            if (!page.IsClosed && !string.Equals(page.Url, beforeUrl, StringComparison.OrdinalIgnoreCase))
            {
                pageTransitionDetected = true;
            }

            if (page.IsClosed && openPages.Count > 0)
            {
                var fallbackPage = openPages[^1];
                return await PrepareNolBookingResultPageAsync(fallbackPage, deadline);
            }

            await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
        }

        if (pageTransitionDetected || (!page.IsClosed && !string.Equals(page.Url, beforeUrl, StringComparison.OrdinalIgnoreCase)))
        {
            var queueSw = Stopwatch.StartNew();
            var lastQueueReportBucket = 0L;
            progress?.Report(new AutomationProgress("대기열 대기 중...", "대기열 진입 — 페이지 전환 대기"));

            while (!cancellationToken.IsCancellationRequested)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var queueReportBucket = (long)(queueSw.Elapsed.TotalSeconds / 10);
                if (queueReportBucket > lastQueueReportBucket)
                {
                    lastQueueReportBucket = queueReportBucket;
                    progress?.Report(new AutomationProgress($"대기열 대기 중... ({(int)queueSw.Elapsed.TotalSeconds}초)"));
                }

                if (page.IsClosed)
                {
                    var remaining = page.Context.Pages.Where(x => !x.IsClosed).ToList();
                    if (remaining.Count > 0)
                        return await PrepareNolBookingResultPageAsync(remaining[^1], DateTimeOffset.UtcNow + timeout);
                    throw new InvalidOperationException("NOL 대기열 페이지가 닫혔습니다.");
                }

                if (await HasNolBookingResultAppearedAsync(page, beforeUrl, beforeTitle))
                {
                    // NOL 대기열 종료 — 예매 페이지 도착
                    return await PrepareNolBookingResultPageAsync(page, DateTimeOffset.UtcNow + timeout);
                }

                var openPages2 = page.Context.Pages.Where(x => !x.IsClosed).ToList();
                var newPage2 = openPages2.FirstOrDefault(x => !beforePages.Contains(x));
                if (newPage2 is not null)
                {
                    // NOL 대기열 종료 — 새 페이지 감지
                    return await PrepareNolBookingResultPageAsync(newPage2, DateTimeOffset.UtcNow + timeout);
                }

                await Task.Delay(100, cancellationToken);
            }
        }

        throw new TimeoutException("NOL 예매하기 클릭 후 페이지 전환을 확인하지 못했습니다.");
    }
    private static async Task<IPage> PrepareNolBookingResultPageAsync(IPage page, DateTimeOffset deadline)
    {
        await page.BringToFrontAsync();
        return page;
    }

    private static async Task<bool> HasNolBookingResultAppearedAsync(IPage page, string beforeUrl, string beforeTitle)
    {
        if (!string.Equals(page.Url, beforeUrl, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var currentTitle = await PlaywrightRuntime.GetPageTitleOrEmptyAsync(page);
        return !string.IsNullOrWhiteSpace(currentTitle) &&
               !string.Equals(currentTitle, beforeTitle, StringComparison.Ordinal) &&
               await page.Locator("#productSide").CountAsync() == 0;
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

        return string.Equals(PlaywrightRuntime.NormalizeText(normalizedActual), PlaywrightRuntime.NormalizeText(normalizedDesired), StringComparison.Ordinal);
    }
    private static bool TryParseNolRound(string value, out string round, out string time)
    {
        round = string.Empty;
        time = string.Empty;
        var normalized = PlaywrightRuntime.NormalizeText(value)
            .Replace("회차", "회", StringComparison.Ordinal)
            .Replace('희', '회')
            .Replace('히', '회')
            .Replace('외', '회');
        var match = NolRoundPattern.Match(normalized);
        if (!match.Success)
        {
            return false;
        }

        round = $"{PlaywrightRuntime.NormalizeText(match.Groups["round"].Value)}회";
        time = NormalizeNolTime(match.Groups["time"].Value);
        return !string.IsNullOrWhiteSpace(round) && !string.IsNullOrWhiteSpace(time);
    }

    private static bool TryParseNolRoundLabel(string value, out string round)
    {
        round = string.Empty;
        var normalized = PlaywrightRuntime.NormalizeText(value).Replace("회차", "회", StringComparison.Ordinal);
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

        round = PlaywrightRuntime.NormalizeText(match.Groups["round"].Value);
        return !string.IsNullOrWhiteSpace(round);
    }

    private static bool TryExtractLeadingRoundNumber(string value, out string roundNumber)
    {
        roundNumber = string.Empty;
        var match = Regex.Match(PlaywrightRuntime.NormalizeText(value), @"^\D*(?<round>\d{1,2})(?:\D|$)", RegexOptions.Compiled);
        if (!match.Success)
        {
            return false;
        }

        roundNumber = PlaywrightRuntime.NormalizeText(match.Groups["round"].Value);
        return !string.IsNullOrWhiteSpace(roundNumber);
    }

    private static string NormalizeNolTime(string value)
    {
        var digits = PlaywrightRuntime.DigitsOnlyPattern.Replace(value, string.Empty);
        if (digits.Length == 3)
        {
            return $"{digits[0]}:{digits[1..]}";
        }

        if (digits.Length == 4)
        {
            return $"{digits[..2]}:{digits[2..]}";
        }

        return PlaywrightRuntime.NormalizeText(value).Replace('.', ':').Replace(',', ':');
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
            var currentTitle = await PlaywrightRuntime.GetPageTitleOrEmptyAsync(page);
            var pageCount = page.Context.Pages.Count;
            var productSideCount = await page.Locator("#productSide").CountAsync();
            var roundButtonCount = await page.Locator(".sideTimeTable .timeTableLabel[role='button']").CountAsync();
            var bookingButtonCount = await page.Locator("#productSide a.sideBtn.is-primary").CountAsync();
            var selectedDateText = await PlaywrightRuntime.GetLocatorTextOrEmptyAsync(page.Locator("#productSide .containerTop .selectedData .date").First);
            var selectedRoundText = await PlaywrightRuntime.GetLocatorTextOrEmptyAsync(page.Locator("#productSide .containerMiddle .selectedData .time").First);
            var openPages = string.Join(", ",
                page.Context.Pages.Select((x, index) => $"[{index}]closed={x.IsClosed};url={PlaywrightRuntime.SafePageUrl(x)}"));
            return $"pageClosed={page.IsClosed}, url={currentUrl}, title={currentTitle}, popupCount={popupCount}, productSideCount={productSideCount}, roundButtonCount={roundButtonCount}, bookingButtonCount={bookingButtonCount}, selectedDate={selectedDateText}, selectedRound={selectedRoundText}, contextPageCount={pageCount}, openPages={openPages}";
        }
        catch (PlaywrightException ex) when (PlaywrightRuntime.IsClosedTargetError(ex))
        {
            return $"page-state-unavailable:{ex.Message}";
        }
    }

}
