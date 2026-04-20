using Microsoft.Extensions.Logging;
using Microsoft.Playwright;
using OpenCvSharp;
using Polly;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using KillRiceMonkey.Application.Abstractions;
using KillRiceMonkey.Application.Models;
using StepTemplate = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.StepTemplate;
using TemplateBounds = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.TemplateBounds;

namespace KillRiceMonkey.Infrastructure.Services;

public sealed class NolAutomationService : INolAutomationService, IAsyncDisposable
{
    private const int NolRemoteDebugPort = 9225;
    private const string NolProfileName = "NolRemoteDebugProfile";
    private const int MaxSeatRetries = 10;
    private const int MaxCaptchaAttempts = 5;
    private const int NolCaptchaLength = 6;
    private const string NolOnestopSeatCircleSelector = "[class*=\"SeatMap_seatGroup\"] circle";
    private const string NolOnestopSeatMapSelector = "[class*='SeatPlan_seatPlan'], [class*='SeatMap_seatMap'], [class*='SeatMap_blockImg'], [class*='SeatMap_placeImg'], [class*='SeatMap_seatGroup']";
    private static readonly Regex NolRoundPattern = new(@"^\D*(?<round>\d{1,2})\s*(?:회차|회|희|히|외)?\s*(?<time>\d{1,2}(?::|\.|,)?\d{2})", RegexOptions.Compiled);
    private static readonly Regex NolSvgNumberPattern = new(@"-?\d+(?:\.\d+)?", RegexOptions.Compiled);
    private const string NolRemoteDebugLaunchUrl = "https://tickets.interpark.com/";
    private const string NolCdpEndpoint = "http://127.0.0.1:9225/";
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
    private static readonly IReadOnlyDictionary<char, char[]> NolCaptchaAlternativeCharMap = new Dictionary<char, char[]>
    {
        ['O'] = ['Q', '0'],
        ['Q'] = ['O', '0'],
        ['0'] = ['O', 'Q'],
        ['I'] = ['L', '1'],
        ['L'] = ['I', '1'],
        ['1'] = ['I', 'L'],
        ['S'] = ['5'],
        ['5'] = ['S'],
        ['B'] = ['8'],
        ['8'] = ['B']
    };
    private readonly ILogger<NolAutomationService> _logger;
    private readonly HttpClient _httpClient = new() { Timeout = TimeSpan.FromSeconds(5) };
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
            NolProfileName);
        Directory.CreateDirectory(profileDirectory);

        var startInfo = new ProcessStartInfo
        {
            FileName = executablePath,
            Arguments = $"--remote-debugging-port={NolRemoteDebugPort} --user-data-dir=\"{profileDirectory}\" --new-window \"{NolRemoteDebugLaunchUrl}\"",
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
        PlaywrightRuntime.EnsureDdddOcrWarmedUp();
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
            var cdpResult = await TryRunNolAutomationViaConnectedBrowserAsync(request, desiredDate, desiredRound, timeout, progress, cancellationToken);
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

    private async Task<AutomationRunResult?> TryRunNolAutomationViaConnectedBrowserAsync(TicketingJobRequest request, DateOnly desiredDate, string desiredRound, TimeSpan timeout, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
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

            if (request.PauseBeforeSeatSelection && request.PauseGate is { } gate)
            {
                _logger.LogInformation("[NOL] 좌석 선택 전 일시정지 — 사용자 재개 대기 중.");
                progress?.Report(new AutomationProgress("좌석 선택 대기 — 일시정지", "좌석 선택 전 일시정지됨. 재개 버튼을 눌러주세요."));
                await Task.Run(() => gate.Wait(cancellationToken), cancellationToken);
                _logger.LogInformation("[NOL] 일시정지 해제 — 좌석 선택 진행.");
                progress?.Report(new AutomationProgress("좌석 선택 중", "일시정지 해제 — 좌석 선택 진행"));
            }

            // 상위 단계용 stage. 하위 SelectNolOnestopSeatAndCompleteAsync/SelectNolLegacySeatAndCompleteAsync가
            // 구역 선택 이후 "좌석 선택 중"을 별도로 report 하므로 여기서는 진입 stage를 구분한다.
            progress?.Report(new AutomationProgress("좌석 페이지 진입", "좌석 선택 페이지로 이동"));
            await SelectNolSeatAndCompleteAsync(captchaPage, timeout, progress, request.DesiredBlock, request.PauseGate, cancellationToken);
            progress?.Report(new AutomationProgress("좌석 선택 완료", "좌석 선택 및 완료 버튼 클릭"));
            return new AutomationRunResult(true, $"NOL 기존 브라우저 DOM 자동화 완료: {desiredDate:yyyy.MM.dd} / {desiredRound} 선택, 좌석 선택 완료.", DateTimeOffset.Now);
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
        var assembly = typeof(NolAutomationService).Assembly;
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
        var excludedCaptchaTexts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var captchaSearchSw = Stopwatch.StartNew();
        _logger.LogInformation("CAPTCHA 입력창 대기 시작 (최대 {Timeout}). url={Url}", timeout, PlaywrightRuntime.SafePageUrl(page));
        var (inputLocator, captchaFrame, foundPage) = await FindNolCaptchaInputAsync(page, timeout, cancellationToken);
        if (foundPage is not null && foundPage != page)
        {
            _logger.LogInformation("[CAPTCHA] 다른 컨텍스트 페이지에서 CAPTCHA 발견. url={Url}, searchMs={Ms}", PlaywrightRuntime.SafePageUrl(foundPage), captchaSearchSw.ElapsedMilliseconds);
            page = foundPage;
        }

        if (inputLocator is null)
        {
            _logger.LogInformation("CAPTCHA 입력창 없음 — CAPTCHA 없는 공연으로 판단, 좌석 선택으로 진행.");
            return;
        }

        _logger.LogInformation("[CAPTCHA] CAPTCHA 입력창 발견 완료. searchMs={Ms}", captchaSearchSw.ElapsedMilliseconds);

        var dialogMessage = string.Empty;
        void OnDialog(object? _, IDialog dialog)
        {
            dialogMessage = dialog.Message;
            _logger.LogInformation("[CAPTCHA] dialog 감지: type={Type}, message={Message}", dialog.Type, dialog.Message);
            dialog.AcceptAsync().ContinueWith(_ => { }, TaskScheduler.Default);
        }

        page.Dialog += OnDialog;
        try
        {
            for (var attempt = 1; attempt <= MaxCaptchaAttempts; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (page.IsClosed) return;

                dialogMessage = string.Empty;
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

                    if (await IsCaptchaGoneAsync(inputLocator, page, captchaFrame))
                    {
                        _logger.LogInformation("[CAPTCHA] CAPTCHA 영역 사라짐 — 통과로 진행. attempt={Attempt}", attempt);
                        return;
                    }

                    if (attempt < MaxCaptchaAttempts)
                        await TryRefreshNolCaptchaImageAsync(page, captchaFrame, cancellationToken);
                    continue;
                }

                if (string.IsNullOrEmpty(text))
                {
                    if (await IsCaptchaGoneAsync(inputLocator, page, captchaFrame))
                    {
                        _logger.LogInformation("[CAPTCHA] OCR 빈 결과 + CAPTCHA 사라짐 — 통과로 진행. attempt={Attempt}", attempt);
                        return;
                    }
                    _logger.LogInformation("[CAPTCHA] OCR 빈 결과 (attempt={Attempt}, {Ms}ms) — 새로고침 후 재시도.", attempt, attemptSw.ElapsedMilliseconds);
                    if (attempt < MaxCaptchaAttempts)
                        await TryRefreshNolCaptchaImageAsync(page, captchaFrame, cancellationToken);
                    continue;
                }

                _logger.LogInformation("CAPTCHA attempt {Attempt}/{Max}: text={Text} ocrMs={OcrMs}", attempt, MaxCaptchaAttempts, text, attemptSw.ElapsedMilliseconds);

                if (text.Length != NolCaptchaLength)
                {
                    _logger.LogInformation("[CAPTCHA] 길이 불일치 ({Len}≠6), 새로고침 후 재시도. attempt={Attempt}", text.Length, attempt);
                    if (attempt < MaxCaptchaAttempts)
                        await TryRefreshNolCaptchaImageAsync(page, captchaFrame, cancellationToken);
                    continue;
                }

                var candidates = BuildNolCaptchaCandidates(text, excludedCaptchaTexts).ToArray();
                if (candidates.Length == 0)
                {
                    _logger.LogInformation("[CAPTCHA] 재시도 가능한 후보 없음 — 새로고침 후 다음 OCR 시도. attempt={Attempt}", attempt);
                    if (attempt < MaxCaptchaAttempts)
                        await TryRefreshNolCaptchaImageAsync(page, captchaFrame, cancellationToken);
                    continue;
                }

                var solved = false;
                foreach (var candidate in candidates)
                {
                    try
                    {
                        await inputLocator.First.FillAsync(candidate, new LocatorFillOptions { Timeout = 300 });
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
                            }", candidate);
                        }
                        catch (PlaywrightException ex)
                        {
                            _logger.LogWarning(ex, "CAPTCHA 입력 실패");
                            excludedCaptchaTexts.Add(candidate);
                            continue;
                        }
                    }

                    const string submitSelector = "button:text-is('입력완료'), a:has-text('입력완료'), a[onclick*='fnCheck']";
                    await ClickNolCaptchaSubmitAsync(page, captchaFrame, inputLocator, submitSelector);
                    _logger.LogInformation("[CAPTCHA] submit 완료. 결과 확인 시작. attempt={Attempt}, candidate={Candidate}", attempt, candidate);

                    var submissionStatus = await GetNolCaptchaSubmissionStatusAsync(page, captchaFrame, inputLocator, cancellationToken);
                    if (!string.IsNullOrEmpty(dialogMessage))
                    {
                        submissionStatus = NolCaptchaSubmissionStatus.Failure;
                        _logger.LogInformation("[CAPTCHA] dialog 감지 — 틀린 CAPTCHA, 재시도. attempt={Attempt}, msg={Msg}", attempt, dialogMessage);
                    }

                    if (submissionStatus == NolCaptchaSubmissionStatus.Success)
                    {
                        _logger.LogInformation("[CAPTCHA] CAPTCHA 제출 완료. attempt={Attempt}, candidate={Candidate}, totalMs={Ms}", attempt, candidate, attemptSw.ElapsedMilliseconds);
                        return;
                    }

                    excludedCaptchaTexts.Add(candidate);
                    _logger.LogInformation("[CAPTCHA] CAPTCHA 실패 감지 — 재시도 대상 제외 추가. attempt={Attempt}, candidate={Candidate}, excluded={Count}", attempt, candidate, excludedCaptchaTexts.Count);
                }

                if (!solved && attempt < MaxCaptchaAttempts)
                    await TryRefreshNolCaptchaImageAsync(page, captchaFrame, cancellationToken);
            }
        }
        finally
        {
            page.Dialog -= OnDialog;
        }

        if (!await IsCaptchaGoneAsync(inputLocator, page, captchaFrame))
            throw new InvalidOperationException($"NOL CAPTCHA 자동 인식 {MaxCaptchaAttempts}회 모두 실패");

        _logger.LogWarning("[CAPTCHA] CAPTCHA 자동 인식 {Max}회 모두 실패했지만 CAPTCHA 영역은 사라졌습니다. 다음 단계로 진행합니다.", MaxCaptchaAttempts);
    }

    private enum NolCaptchaSubmissionStatus
    {
        Success,
        Failure
    }

    private static IEnumerable<string> BuildNolCaptchaCandidates(string recognizedText, HashSet<string> excludedCaptchaTexts)
    {
        var ordered = new List<string>();
        void AddCandidate(string candidate)
        {
            if (candidate.Length != NolCaptchaLength)
                return;

            candidate = candidate.ToUpperInvariant();
            if (excludedCaptchaTexts.Contains(candidate) || ordered.Contains(candidate, StringComparer.OrdinalIgnoreCase))
                return;

            ordered.Add(candidate);
        }

        AddCandidate(recognizedText);
        foreach (var variant in EnumerateNolCaptchaVariants(recognizedText.ToUpperInvariant(), 0, 0))
            AddCandidate(variant);

        return ordered;
    }

    private static IEnumerable<string> EnumerateNolCaptchaVariants(string text, int index, int substitutions)
    {
        if (index >= text.Length || substitutions >= 2)
            yield break;

        for (var i = index; i < text.Length; i++)
        {
            if (!NolCaptchaAlternativeCharMap.TryGetValue(text[i], out var alternatives))
                continue;

            foreach (var alternative in alternatives)
            {
                var chars = text.ToCharArray();
                chars[i] = alternative;
                var variant = new string(chars);
                yield return variant;

                foreach (var child in EnumerateNolCaptchaVariants(variant, i + 1, substitutions + 1))
                    yield return child;
            }
        }
    }

    private async Task<NolCaptchaSubmissionStatus> GetNolCaptchaSubmissionStatusAsync(IPage page, IFrame? captchaFrame, ILocator inputLocator, CancellationToken cancellationToken)
    {
        // 티켓팅 특성상 한 번의 제출 판정이 오래 걸릴수록 전체 재시도 시간이 크게 늘어난다.
        // 초기 지연을 최소화(50ms)하고 폴링 간격도 25ms로 좁혀 빠르게 성공/실패를 판정한다.
        // deadline은 600ms로 단축 — 서버가 이 시간 안에 성공/실패 응답을 반영하지 못하면 이후에도
        // 직후 단계에서 최종 판정 로직이 한 번 더 체크하므로 오탐을 높이지 않는다.
        await Task.Delay(50, cancellationToken);

        var deadline = DateTimeOffset.UtcNow + TimeSpan.FromMilliseconds(600);
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            // Failure 신호가 우선. gone 체크가 오탐(잠깐 숨겨진 에러 div 등)하는 현상 방지.
            if (await HasCaptchaInlineErrorAsync(page, captchaFrame))
                return NolCaptchaSubmissionStatus.Failure;

            if (await IsCaptchaGoneAsync(inputLocator, page, captchaFrame))
                return NolCaptchaSubmissionStatus.Success;

            await Task.Delay(25, cancellationToken);
        }

        // 최종: 에러가 확정되면 Failure, 그 외에 모달이 사라져 있으면 Success
        if (await HasCaptchaInlineErrorAsync(page, captchaFrame))
            return NolCaptchaSubmissionStatus.Failure;

        return await IsCaptchaGoneAsync(inputLocator, page, captchaFrame)
            ? NolCaptchaSubmissionStatus.Success
            : NolCaptchaSubmissionStatus.Failure;
    }

    private static async Task ClickNolCaptchaSubmitAsync(IPage page, IFrame? captchaFrame, ILocator inputLocator, string submitSelector)
    {
        const string directClickScript = @"() => {
            const button = Array.from(document.querySelectorAll('button, a')).find(el => (el.innerText || '').trim() === '입력완료');
            if (!button) return false;
            if ('disabled' in button) button.disabled = false;
            if (typeof button.click === 'function') button.click();
            else button.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
            return true;
        }";

        try
        {
            var clicked = captchaFrame is not null
                ? await captchaFrame.EvaluateAsync<bool>(directClickScript)
                : await page.EvaluateAsync<bool>(directClickScript);
            if (clicked)
                return;
        }
        catch (PlaywrightException) { }

        ILocator? submitLocator = null;

        if (captchaFrame is not null)
        {
            var frameSubmit = captchaFrame.Locator(submitSelector);
            try { if (await frameSubmit.CountAsync() > 0) submitLocator = frameSubmit; } catch { }
        }

        if (submitLocator is null)
        {
            var pageSubmit = page.Locator(submitSelector);
            try { if (await pageSubmit.CountAsync() > 0) submitLocator = pageSubmit; } catch { }
        }

        if (submitLocator is not null)
        {
            try
            {
                await submitLocator.First.EvaluateAsync("el => { if (el.disabled) el.disabled = false; el.click(); }");
                return;
            }
            catch (PlaywrightException)
            {
                try { await submitLocator.First.ClickAsync(new LocatorClickOptions { Timeout = 500, Force = true }); return; }
                catch (PlaywrightException) { }
            }
        }

        try
        {
            await inputLocator.First.EvaluateAsync(@"el => {
                if (typeof fnCheck === 'function') { fnCheck(); }
                else if (el.form) { el.form.submit(); }
                else { el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, bubbles: true })); }
            }");
        }
        catch (PlaywrightException)
        {
            try { await inputLocator.First.PressAsync("Enter"); } catch (PlaywrightException) { }
        }
    }

    private static async Task<bool> IsCaptchaGoneAsync(ILocator inputLocator, IPage page, IFrame? captchaFrame)
    {
        try
        {
            if (await inputLocator.CountAsync() == 0) return true;
        }
        catch { return true; }

        try
        {
            var hidden = captchaFrame is not null
                ? await captchaFrame.EvaluateAsync<bool>(@"() => {
                    const el = document.querySelector('#divCaptchaWrap, #divCaptcha_R, .captcha_area, .wrap_captcha');
                    if (!el) return false;
                    const s = window.getComputedStyle(el);
                    return s.display === 'none' || s.visibility === 'hidden' || s.opacity === '0';
                  }")
                : await page.EvaluateAsync<bool>(@"() => {
                    // NOL NEW 기준: 실제 모달 컨테이너(ModalCaptchaText_layerWrap)만 판정.
                    // [class*='ModalCaptchaText']로만 잡으면 ModalCaptchaText_captchaError 같은
                    // 항상 존재하는 내부 요소가 매칭되어 gone 오탐이 발생한다.
                    const modal = document.querySelector('[class*=""ModalCaptchaText_layerWrap""], [class*=""captchaModal""], .captcha_area, .wrap_captcha, #divCaptchaWrap, #divCaptcha_R');
                    if (!modal) return true;
                    const s = window.getComputedStyle(modal);
                    return s.display === 'none' || s.visibility === 'hidden' || s.opacity === '0';
                  }");
            if (hidden) return true;
        }
        catch { }

        return false;
    }

    private static async Task<bool> HasCaptchaInlineErrorAsync(IPage page, IFrame? captchaFrame)
    {
        // 실측 결과 실패 시 DOM 상태:
        //   1) <div class="ModalCaptchaText_captchaError__...">입력한 문자를 다시 확인해주세요</div>
        //   2) <input class="ModalCaptchaText_captchaInput__... ModalCaptchaText_invalid__..."> (invalid 클래스 부착)
        //   3) ModalCaptchaText_captchaContent 내부 innerText에 오류 문구 포함
        // 주의: captchaError/invalid 클래스가 빈 상태로 DOM에 미리 존재할 수 있으므로
        //       클래스 존재 + innerText 비어있지 않음 조합으로 검증한다.
        const string script = @"() => {
            // 1) captchaError 요소의 텍스트가 실제로 채워진 경우
            const errDiv = document.querySelector('[class*=""ModalCaptchaText_captchaError""], [class*=""captchaError""]');
            if (errDiv) {
                const t = (errDiv.innerText || errDiv.textContent || '').trim();
                if (t.length > 0) return true;
            }

            // 2) 입력창에 invalid 클래스가 붙은 경우
            if (document.querySelector('input[class*=""ModalCaptchaText_invalid""], input[class*=""captchaInput""][class*=""invalid""]'))
                return true;

            // 3) 모달/컨테이너 innerText에 구체적 에러 문구가 포함된 경우
            const containers = document.querySelectorAll('[class*=""ModalCaptchaText_layerWrap""], [class*=""ModalCaptchaText_captchaContent""], [class*=""captchaModal""], #divCaptchaWrap, #divCaptcha_R, .captcha_area, .wrap_captcha');
            for (const c of containers) {
                const t = c.innerText || '';
                if (t.includes('입력한 문자') || t.includes('다시 확인') || t.includes('일치하지') || t.includes('올바른 문자') || t.includes('정확하게 입력'))
                    return true;
            }
            return false;
        }";

        try
        {
            return captchaFrame is not null
                ? await captchaFrame.EvaluateAsync<bool>(script)
                : await page.EvaluateAsync<bool>(script);
        }
        catch { return false; }
    }

    private async Task TryRefreshNolCaptchaImageAsync(IPage page, IFrame? captchaFrame, CancellationToken cancellationToken)
    {
        if (page.IsClosed)
            return;

        var previousSrc = await TryGetCurrentNolCaptchaImageSourceAsync(page, captchaFrame);

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
                await WaitForNolCaptchaImageRefreshAsync(page, captchaFrame, previousSrc, cancellationToken);
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
                await WaitForNolCaptchaImageRefreshAsync(page, captchaFrame, previousSrc, cancellationToken);
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
            await WaitForNolCaptchaImageRefreshAsync(page, captchaFrame, previousSrc, cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "CAPTCHA 새로고침 모든 방법 실패");
        }
    }

    private static async Task<string?> TryGetCurrentNolCaptchaImageSourceAsync(IPage page, IFrame? captchaFrame)
    {
        const string script = @"() => document.querySelector('img[alt=""캡챠 이미지""], #imgCaptcha, #captchaImg, img[src*=""captcha"" i], img[src*=""cap_img"" i]')?.getAttribute('src') ?? null";
        try
        {
            return captchaFrame is not null
                ? await captchaFrame.EvaluateAsync<string?>(script)
                : await page.EvaluateAsync<string?>(script);
        }
        catch
        {
            return null;
        }
    }

    private static async Task WaitForNolCaptchaImageRefreshAsync(IPage page, IFrame? captchaFrame, string? previousSrc, CancellationToken cancellationToken)
    {
        var deadline = DateTimeOffset.UtcNow + TimeSpan.FromMilliseconds(500);
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var currentSrc = await TryGetCurrentNolCaptchaImageSourceAsync(page, captchaFrame);
            if (!string.IsNullOrWhiteSpace(currentSrc) && !string.Equals(currentSrc, previousSrc, StringComparison.Ordinal))
                return;

            await Task.Delay(50, cancellationToken);
        }
    }

    private static async Task<(ILocator? inputLocator, IFrame? frame, IPage? foundPage)> FindNolCaptchaInputAsync(
        IPage page, TimeSpan timeout, CancellationToken cancellationToken)
    {
        const string inputSelector = "#txtCaptcha, [class*='captchaInput'] input, [class*='captchaInput'], input[placeholder*='문자'], input[name*='captcha' i], input[id*='captcha' i], input[name*='CAPTCHA'], input[placeholder*='보안문자'], input[placeholder*='자동입력']";
        var infinite = timeout == Timeout.InfiniteTimeSpan;
        var timeoutMs = infinite ? int.MaxValue : (float)timeout.TotalMilliseconds;

        using var innerCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var winnerTcs = new TaskCompletionSource<(ILocator?, IFrame?, IPage?)>(TaskCreationOptions.RunContinuationsAsynchronously);

        // 이벤트 기반: 각 context page + 주요 frame에 대해 Locator.WaitForAsync를 병렬 발사.
        // 서버 응답이 도착해 DOM에 입력창이 attach되는 즉시 성공 반환 → 폴링 30ms overhead 제거.
        var waitTasks = new List<Task>();

        void RegisterPage(IPage targetPage)
        {
            if (targetPage.IsClosed) return;

            // main frame
            waitTasks.Add(WaitOneAsync(targetPage, null, inputSelector, timeoutMs, innerCts.Token, winnerTcs));
            // 현재 존재하는 child frames
            try
            {
                foreach (var f in targetPage.Frames)
                {
                    if (f == targetPage.MainFrame) continue;
                    waitTasks.Add(WaitOneAsync(targetPage, f, inputSelector, timeoutMs, innerCts.Token, winnerTcs));
                }
            }
            catch (PlaywrightException) { }

            // 이후 등장하는 새 frame도 감지 (OLD iframe 새로 붙을 수 있음)
            targetPage.FrameAttached += (_, frame) =>
            {
                if (innerCts.IsCancellationRequested) return;
                if (frame == targetPage.MainFrame) return;
                waitTasks.Add(WaitOneAsync(targetPage, frame, inputSelector, timeoutMs, innerCts.Token, winnerTcs));
            };
        }

        try
        {
            RegisterPage(page);
            foreach (var p in page.Context.Pages)
            {
                if (p != page) RegisterPage(p);
            }

            // 새 페이지 열릴 경우도 감지
            page.Context.Page += (_, newPage) =>
            {
                if (innerCts.IsCancellationRequested) return;
                if (newPage != page) RegisterPage(newPage);
            };

            if (waitTasks.Count == 0)
                return (null, null, null);

            // winner 또는 전체 timeout 대기
            var timeoutTask = infinite ? Task.Delay(Timeout.InfiniteTimeSpan, innerCts.Token)
                                       : Task.Delay(timeout, innerCts.Token);
            var completed = await Task.WhenAny(winnerTcs.Task, timeoutTask);
            if (completed == winnerTcs.Task)
                return await winnerTcs.Task;
            return (null, null, null);
        }
        finally
        {
            // 대기 중인 WaitForAsync들을 취소해 playwright 부하 해소
            innerCts.Cancel();
            try { await Task.WhenAll(waitTasks).WaitAsync(TimeSpan.FromMilliseconds(50)); } catch { }
        }

        static async Task WaitOneAsync(
            IPage owningPage, IFrame? frame, string selector, float timeoutMs,
            CancellationToken token, TaskCompletionSource<(ILocator?, IFrame?, IPage?)> winner)
        {
            try
            {
                var baseLocator = frame is null ? owningPage.Locator(selector) : frame.Locator(selector);
                await baseLocator.First.WaitForAsync(new LocatorWaitForOptions
                {
                    State = WaitForSelectorState.Attached,
                    Timeout = timeoutMs,
                }).WaitAsync(token);

                if (token.IsCancellationRequested) return;
                winner.TrySetResult((baseLocator, frame, owningPage));
            }
            catch (OperationCanceledException) { }
            catch (PlaywrightException) { }
            catch (Exception) { }
        }
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
        IPage? queueCandidatePage = null;

        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var openPages = page.Context.Pages.Where(x => !x.IsClosed).ToList();
            var newPage = openPages.FirstOrDefault(x => !beforePages.Contains(x));
            if (newPage is not null)
            {
                if (await HasNolBookingResultAppearedAsync(newPage, string.Empty, string.Empty))
                    return await PrepareNolBookingResultPageAsync(newPage, deadline);

                queueCandidatePage ??= newPage;
            }

            if (!page.IsClosed && await HasNolBookingResultAppearedAsync(page, beforeUrl, beforeTitle))
            {
                return await PrepareNolBookingResultPageAsync(page, deadline);
            }

            if (!page.IsClosed &&
                !string.Equals(StripUrlFragment(page.Url), StripUrlFragment(beforeUrl), StringComparison.OrdinalIgnoreCase))
            {
                pageTransitionDetected = true;
                queueCandidatePage ??= page;
            }

            if (page.IsClosed && openPages.Count > 0)
            {
                var fallbackPage = openPages[^1];
                return await PrepareNolBookingResultPageAsync(fallbackPage, deadline);
            }

            await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
        }

        if (queueCandidatePage is not null ||
            pageTransitionDetected ||
            (!page.IsClosed && !string.Equals(StripUrlFragment(page.Url), StripUrlFragment(beforeUrl), StringComparison.OrdinalIgnoreCase)))
        {
            var queueSw = Stopwatch.StartNew();
            var lastQueueReportBucket = 0L;
            progress?.Report(new AutomationProgress("대기열 대기 중...", "대기열 진입 — 페이지 전환 대기"));
            var queuePage = queueCandidatePage ?? page;

            while (!cancellationToken.IsCancellationRequested)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var queueReportBucket = (long)(queueSw.Elapsed.TotalSeconds / 10);
                if (queueReportBucket > lastQueueReportBucket)
                {
                    lastQueueReportBucket = queueReportBucket;
                    progress?.Report(new AutomationProgress($"대기열 대기 중... ({(int)queueSw.Elapsed.TotalSeconds}초)"));
                }

                if (queuePage.IsClosed)
                {
                    var remaining = page.Context.Pages.Where(x => !x.IsClosed).ToList();
                    if (remaining.Count > 0)
                        return await PrepareNolBookingResultPageAsync(remaining[^1], DateTimeOffset.UtcNow + timeout);
                    throw new InvalidOperationException("NOL 대기열 페이지가 닫혔습니다.");
                }

                if (await HasNolBookingResultAppearedAsync(queuePage, beforeUrl, beforeTitle))
                {
                    return await PrepareNolBookingResultPageAsync(queuePage, DateTimeOffset.UtcNow + timeout);
                }

                var openPages2 = page.Context.Pages.Where(x => !x.IsClosed).ToList();
                var newPage2 = openPages2.FirstOrDefault(x => !beforePages.Contains(x));
                if (newPage2 is not null)
                {
                    if (await HasNolBookingResultAppearedAsync(newPage2, string.Empty, string.Empty))
                    {
                        return await PrepareNolBookingResultPageAsync(newPage2, DateTimeOffset.UtcNow + timeout);
                    }

                    queuePage = newPage2;
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
        try
        {
            var currentUrl = page.Url;
            var currentTitle = await PlaywrightRuntime.GetPageTitleOrEmptyAsync(page);
            var hasProductSide = await page.Locator("#productSide").CountAsync() > 0;
            var hasCaptchaInput = await page.Locator("#txtCaptcha, [class*='captchaInput'] input, [class*='ModalCaptchaText_captchaInput'] input, input[placeholder*='문자'], input[name*='captcha' i], input[id*='captcha' i], input[name*='CAPTCHA'], input[placeholder*='보안문자'], input[placeholder*='자동입력']").CountAsync() > 0;
            var hasCaptchaModal = await page.Locator("[class*='ModalCaptchaText_layerWrap'], [class*='captchaModal'], #divCaptchaWrap, #divCaptcha_R, .captcha_area, .wrap_captcha").CountAsync() > 0;
            var hasLegacySeatFrame = page.Frames.Any(f => f != page.MainFrame && string.Equals(f.Name, "ifrmSeat", StringComparison.OrdinalIgnoreCase));
            var hasOnestopSeatMap = await page.Locator("svg circle, [class*='SeatPlan_seatPlan'], [class*='SeatMap_seatGroup']").CountAsync() > 0;

            return NolBookingPageClassifier.Classify(
                       currentUrl,
                       beforeUrl,
                       currentTitle,
                       beforeTitle,
                       hasProductSide,
                       hasCaptchaInput,
                       hasCaptchaModal,
                       hasLegacySeatFrame,
                       hasOnestopSeatMap) == NolBookingPageState.BookingReady;
        }
        catch (PlaywrightException ex) when (PlaywrightRuntime.IsClosedTargetError(ex))
        {
            return true;
        }
    }

    private async Task SelectNolSeatAndCompleteAsync(IPage captchaPage, TimeSpan timeout, IProgress<AutomationProgress>? progress, string? desiredBlock, ManualResetEventSlim? pauseGate, CancellationToken cancellationToken)
    {
        var totalSw = Stopwatch.StartNew();
        var (seatPage, isOnestop) = await FindNolSeatPageAsync(captchaPage, timeout, cancellationToken);
        _logger.LogInformation("[SeatSelect] 좌석 선택 페이지 감지. isOnestop={IsOnestop}, url={Url}", isOnestop, PlaywrightRuntime.SafePageUrl(seatPage));

        await seatPage.BringToFrontAsync();

        var excludedSeats = new HashSet<string>();

        for (var attempt = 0; attempt < MaxSeatRetries; attempt++)
        {
            try
            {
                if (isOnestop)
                    await SelectNolOnestopSeatAndCompleteAsync(seatPage, timeout, progress, desiredBlock, excludedSeats, pauseGate, cancellationToken);
                else
                    await SelectNolLegacySeatAndCompleteAsync(seatPage, timeout, progress, desiredBlock, excludedSeats, pauseGate, cancellationToken);

                _logger.LogInformation("[SeatSelect] 좌석 선택 완료. totalMs={Ms}", totalSw.ElapsedMilliseconds);
                return;
            }
            catch (Exception ex) when (attempt < MaxSeatRetries - 1)
            {
                _logger.LogWarning(ex, "[SeatSelect] 좌석 선택 실패 (attempt={Attempt}/{Max}). 재시도.", attempt + 1, MaxSeatRetries);
                if (seatPage.IsClosed)
                {
                    _logger.LogWarning("[SeatSelect] 좌석 페이지가 닫혀 있어 재시도 불가.");
                    break;
                }
                await Task.Delay(100, cancellationToken);
            }
        }

        throw new InvalidOperationException($"좌석 선택 {MaxSeatRetries}회 시도 모두 실패.");
    }

    private static async Task<(IPage seatPage, bool isOnestop)> FindNolSeatPageAsync(IPage contextPage, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var deadline = DateTimeOffset.UtcNow + timeout;

        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            foreach (var p in contextPage.Context.Pages.Where(x => !x.IsClosed))
            {
                var url = PlaywrightRuntime.SafePageUrl(p);
                if (url.Contains("/onestop/seat", StringComparison.OrdinalIgnoreCase) ||
                    url.Contains("/onestop/", StringComparison.OrdinalIgnoreCase))
                    return (p, true);
                if (url.Contains("poticket.interpark.com", StringComparison.OrdinalIgnoreCase))
                    return (p, false);
            }

            await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
        }

        return (contextPage, !contextPage.Frames.Any(f => f != contextPage.MainFrame &&
            f.Url.Contains("poticket", StringComparison.OrdinalIgnoreCase)));
    }

    private async Task SelectNolOnestopSeatAndCompleteAsync(IPage page, TimeSpan timeout, IProgress<AutomationProgress>? progress, string? desiredBlock, HashSet<string> excludedSeats, ManualResetEventSlim? pauseGate, CancellationToken cancellationToken)
    {
        var stepSw = Stopwatch.StartNew();

        await WaitForNolOnestopSeatMapAsync(page, timeout, cancellationToken);
        _logger.LogInformation("[OnestopSeat] 좌석맵 로드 완료. waitMs={Ms}", stepSw.ElapsedMilliseconds);

        stepSw.Restart();
        var zoneRequired = await IsNolOnestopZoneSelectionRequiredAsync(page);
        _logger.LogInformation("[OnestopSeat] 구역 선택 필요={Required}. checkMs={Ms}", zoneRequired, stepSw.ElapsedMilliseconds);

        NolZoneSvgCandidate? selectedZone = null;
        IReadOnlyList<NolZoneSvgCandidate>? allZones = null;
        if (zoneRequired)
        {
            stepSw.Restart();
            progress?.Report(new AutomationProgress("구역 선택 중", "구역 자동 선택 시도 중"));
            (selectedZone, allZones) = await EnsureNolOnestopZoneSelectedAsync(page, desiredBlock, progress, cancellationToken);
            _logger.LogInformation("[OnestopSeat] 구역 선택 완료. fill={Fill}, bounds=({X:F1},{Y:F1},{W:F1},{H:F1}), allZones={AllCount}, waitMs={Ms}",
                selectedZone?.Fill ?? "<unknown>",
                selectedZone?.X ?? double.NaN, selectedZone?.Y ?? double.NaN, selectedZone?.Width ?? double.NaN, selectedZone?.Height ?? double.NaN,
                allZones?.Count ?? 0,
                stepSw.ElapsedMilliseconds);
        }
        var selectedZoneFill = selectedZone?.Fill.ToLowerInvariant();

        stepSw.Restart();
        progress?.Report(new AutomationProgress("좌석 선택 중"));
        // NOTE: 구역 선택 직후 TryResetNolSeatPlanZoomAsync(좌석도 전체보기)를 호출하면
        // 해당 버튼이 구역 선택 상태를 해제하고 구역 선택 화면으로 되돌려버려
        // 사용자가 수동으로 선택한 구역이 무한 반복으로 해제되는 버그가 발생한다.
        // 좌석 클릭은 아래 SelectNolOnestopSeatAsync에서 ScrollIntoViewIfNeededAsync +
        // Locator.ClickAsync(force)로 Playwright가 자동 스크롤/대기를 처리하도록 한다.
        // selectedZone이 지정되면 해당 구역의 fill(색) + bounds(좌표 범위) 둘 다로 좌석을 제한한다.
        // NEW URL의 좌석맵은 선택하지 않은 다른 구역의 좌석도 disabled 없이 렌더링되며,
        // 같은 fill(예: 주황)의 여러 지정석 구역(A/B/G/H ...)이 모두 함께 활성화된다.
        // 추가로 큰 구역(G/H) bounds 안에 작은 구역(A1/A2/A3)이 물리적으로 포함되어 단순 bounds 검사로는
        // 큰 구역을 선택해도 작은 구역 좌석이 후보에 남는다. "exclusive zone membership"(좌석을
        // 포함하는 모든 candidate 중 가장 작은 bbox가 selectedZone인 경우에만 후보 허용)으로 완전 해결.
        var selectedSeatId = await SelectNolOnestopSeatAsync(page, excludedSeats, selectedZone, allZones, pauseGate, progress, cancellationToken);
        _logger.LogInformation("[OnestopSeat] 좌석 클릭 완료. seatId={SeatId}, selectMs={Ms}", selectedSeatId, stepSw.ElapsedMilliseconds);

        stepSw.Restart();
        await ClickNolOnestopSeatCompleteAsync(page, timeout, cancellationToken);
        _logger.LogInformation("[OnestopSeat] 선택 완료 버튼 클릭. completeMs={Ms}", stepSw.ElapsedMilliseconds);
    }

    private static async Task WaitForNolOnestopSeatMapAsync(IPage page, TimeSpan timeout, CancellationToken cancellationToken)
    {
        await PlaywrightRuntime.WaitForConditionAsync(
            async () =>
            {
                try
                {
                    if (await page.Locator(NolOnestopSeatMapSelector).CountAsync() > 0)
                        return true;

                    var circleCount = await page.Locator(NolOnestopSeatCircleSelector).CountAsync();
                    return circleCount > 0;
                }
                catch (PlaywrightException) { return false; }
            },
            timeout,
            cancellationToken,
            "NOL 원스탑 좌석/구역 맵을 찾지 못했습니다.");
    }

    private static async Task<bool> IsNolOnestopZoneSelectionRequiredAsync(IPage page)
    {
        try
        {
            var available = await page.EvaluateAsync<int>($@"() => {{
                const circles = document.querySelectorAll('{NolOnestopSeatCircleSelector}');
                let count = 0;
                for (const c of circles) {{
                    const cls = c.getAttribute('class') || '';
                    if (cls.includes('disabled')) continue;
                    count++;
                }}
                return count;
            }}");
            return available < 10;
        }
        catch (PlaywrightException) { return false; }
    }

    private sealed record NolOnestopZoneCandidate(string Key, double ClientX, double ClientY, string Fill, double Area);

    private async Task<(NolZoneSvgCandidate? Selected, IReadOnlyList<NolZoneSvgCandidate>? AllCandidates)> EnsureNolOnestopZoneSelectedAsync(IPage page, string? desiredBlock, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        // 반환값: (선택된 구역 candidate, SVG 전체 후보 목록).
        // AllCandidates가 있으면 SelectNolOnestopSeatAsync가 "exclusive zone membership" 필터를 적용해
        // 큰 구역(G/H) 안에 포함된 작은 구역(A1/A2/A3) 좌석이 큰 구역 후보에 섞이는 버그를 차단할 수 있다.
        //
        // DesiredBlock이 비어 있으면 구역 자동 선택을 건너뛰고 사용자가 수동으로 구역을 클릭할 때까지 대기한다.
        // 좌석(SeatMap_seatGroup)에 활성 circle이 충분히 로드되면 사용자 선택 완료로 간주.
        if (string.IsNullOrWhiteSpace(desiredBlock))
        {
            _logger.LogInformation("[OnestopSeat] DesiredBlock 비어있음 — 구역 수동 선택 대기 모드.");
            progress?.Report(new AutomationProgress("구역 선택 대기 중", "구역을 직접 클릭해 주세요 (DesiredBlock 미지정)."));

            // 사용자의 mousedown 좌표를 JS에서 캡처. 구역 선택 완료 후 blockImg이 사라지므로
            // "클릭 당시" 좌표/이미지 rect/SVG URL 을 미리 기록한다.
            await InstallUserZoneClickListenerAsync(page);

            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var available = await CountAvailableNolOnestopSeatsAsync(page, null);
                if (available >= 10)
                {
                    // 1) 사용자 클릭 좌표를 이용해 SVG 후보 중 포함되는 구역 찾기 → 후보 + 전체 목록 반환.
                    var (preciseCand, allCands) = await ResolveUserClickZoneCandidateWithAllAsync(page);
                    if (preciseCand is not null)
                    {
                        _logger.LogInformation("[OnestopSeat] 사용자 구역 선택 감지 — 좌석 {Count}개 로드, 클릭→fill={Fill}, bounds=({X:F1},{Y:F1},{W:F1},{H:F1}), allZones={AllCount}.",
                            available, preciseCand.Fill, preciseCand.X, preciseCand.Y, preciseCand.Width, preciseCand.Height, allCands?.Count ?? 0);
                        return (preciseCand, allCands);
                    }

                    // 2) fallback: dominant fill 만 사용(bounds 없이 fill 필터만). 정확도 하락 경고.
                    var dominantFill = await DetectDominantActiveSeatFillAsync(page);
                    _logger.LogWarning("[OnestopSeat] 사용자 클릭 좌표 매칭 실패 — dominant fill fallback={Fill}, 좌석 {Count}개. fill 필터만 적용(bounds 미지정).", dominantFill ?? "<none>", available);
                    var fallback = dominantFill is null
                        ? null
                        : new NolZoneSvgCandidate(Key: "<fill-only>", Fill: dominantFill, Area: 0, X: double.NaN, Y: double.NaN, Width: double.NaN, Height: double.NaN, ViewWidth: 0, ViewHeight: 0);
                    return (fallback, null);
                }

                // 25ms 폴링: 사람 반응속도(~150ms) 대비 6배 빠르게 감지. CPU 부담 미미.
                await Task.Delay(25, cancellationToken);
            }
        }

        _logger.LogInformation("[OnestopSeat] 구역 자동 선택 시작. desiredBlock={Desired}", desiredBlock);

        var attemptedZones = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var deadline = DateTimeOffset.UtcNow + TimeSpan.FromSeconds(15);
        var clearCount = 0;
        const int MaxClearCount = 3;
        NolOnestopZoneCandidate? lastClicked = null;
        NolZoneSvgCandidate? lastSvgCandidate = null;
        IReadOnlyList<NolZoneSvgCandidate>? lastAllCandidates = null;

        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                var available = await CountAvailableNolOnestopSeatsAsync(page, lastClicked?.Fill.ToLowerInvariant());
                if (available >= 10)
                {
                    _logger.LogInformation("[OnestopSeat] 구역 선택 후 좌석 {Count}개 감지(fill={Fill}).", available, lastClicked?.Fill ?? "<any>");
                    return (lastSvgCandidate, lastAllCandidates);
                }

                var (zoneCandidate, svgCandidate, allCandidates) = await TryFindNolOnestopZoneCandidateTripleAsync(page, desiredBlock, attemptedZones);
                if (zoneCandidate is not null)
                {
                    attemptedZones.Add(zoneCandidate.Key);

                    var viewport = await GetNolViewportSizeAsync(page);
                    if (!IsViewportPoint(zoneCandidate.ClientX, zoneCandidate.ClientY, viewport))
                    {
                        _logger.LogInformation("[OnestopSeat] 후보 좌표 viewport 밖 — 다음 후보로 진행. key={Key}, point=({X:F1},{Y:F1}), viewport=({VW},{VH})",
                            zoneCandidate.Key, zoneCandidate.ClientX, zoneCandidate.ClientY, viewport.width, viewport.height);
                        continue;
                    }

                    await page.Mouse.MoveAsync((float)zoneCandidate.ClientX, (float)zoneCandidate.ClientY);
                    await page.Mouse.DownAsync();
                    await page.Mouse.UpAsync();
                    lastClicked = zoneCandidate;
                    lastSvgCandidate = svgCandidate;
                    lastAllCandidates = allCandidates;
                    _logger.LogInformation("[OnestopSeat] 구역 후보 클릭 — key={Key}, fill={Fill}, area={Area:F0}, x={X:F1}, y={Y:F1}",
                        zoneCandidate.Key, zoneCandidate.Fill, zoneCandidate.Area, zoneCandidate.ClientX, zoneCandidate.ClientY);

                    var loaded = await WaitForNolOnestopSeatsLoadedAsync(page, zoneCandidate.Fill.ToLowerInvariant(), TimeSpan.FromMilliseconds(1500), cancellationToken);
                    if (loaded)
                    {
                        _logger.LogInformation("[OnestopSeat] 구역 클릭 후 좌석 로드 확인(fill={Fill}, bounds=({X:F1},{Y:F1},{W:F1},{H:F1}), allZones={AllCount}).",
                            zoneCandidate.Fill,
                            svgCandidate?.X ?? double.NaN, svgCandidate?.Y ?? double.NaN, svgCandidate?.Width ?? double.NaN, svgCandidate?.Height ?? double.NaN,
                            allCandidates?.Count ?? 0);
                        return (svgCandidate, allCandidates);
                    }
                    continue;
                }

                if (attemptedZones.Count > 0)
                {
                    clearCount++;
                    if (clearCount >= MaxClearCount)
                    {
                        throw new InvalidOperationException($"NOL 구역 자동 선택 실패 — 후보 {attemptedZones.Count}개 시도 모두 좌석 없음 (clear {MaxClearCount}회).");
                    }

                    _logger.LogInformation("[OnestopSeat] 시도한 구역({Count})이 모두 소진 — 초기화 후 재탐색 ({Clear}/{Max}).",
                        attemptedZones.Count, clearCount, MaxClearCount);
                    attemptedZones.Clear();
                }
            }
            catch (PlaywrightException ex)
            {
                _logger.LogDebug(ex, "[OnestopSeat] 구역 선택 루프 중 일시 오류 — 재시도.");
            }

            // 자동 구역 재시도 폴링 (빠르게 다음 iteration)
            await Task.Delay(30, cancellationToken);
        }

        throw new TimeoutException("NOL 구역 자동 선택 시간 초과 (15초).");
    }

    // ResolveUserClickZoneCandidateAsync의 확장판: candidate + 전체 목록 함께 반환.
    private async Task<(NolZoneSvgCandidate? Selected, IReadOnlyList<NolZoneSvgCandidate>? All)> ResolveUserClickZoneCandidateWithAllAsync(IPage page)
    {
        try
        {
            var clickJson = await page.EvaluateAsync<string?>(@"() => {
                const c = window.__nolUserClick;
                return c ? JSON.stringify(c) : null;
            }");
            if (string.IsNullOrEmpty(clickJson))
                return (null, null);

            using var doc = JsonDocument.Parse(clickJson);
            var root = doc.RootElement;
            var clickX = root.GetProperty("x").GetDouble();
            var clickY = root.GetProperty("y").GetDouble();
            var imgLeft = root.GetProperty("imgLeft").GetDouble();
            var imgTop = root.GetProperty("imgTop").GetDouble();
            var imgW = root.GetProperty("imgW").GetDouble();
            var imgH = root.GetProperty("imgH").GetDouble();
            var svgSrc = root.GetProperty("svgSrc").GetString();

            if (imgW <= 0 || imgH <= 0 || string.IsNullOrEmpty(svgSrc))
                return (null, null);

            var svgText = await _httpClient.GetStringAsync(svgSrc);
            var all = BuildNolOnestopZoneCandidatesFromSvg(svgText).ToList();
            if (all.Count == 0)
                return (null, null);

            var vw = all[0].ViewWidth;
            var vh = all[0].ViewHeight;
            if (vw <= 0 || vh <= 0)
                return (null, null);

            var svgX = (clickX - imgLeft) / imgW * vw;
            var svgY = (clickY - imgTop) / imgH * vh;

            var containing = all
                .Where(c => c.ContainsSvgPoint(svgX, svgY))
                .OrderBy(c => (c.CenterX - svgX) * (c.CenterX - svgX) + (c.CenterY - svgY) * (c.CenterY - svgY))
                .ToList();
            if (containing.Count > 0)
                return (containing[0], all);

            var tolerance = Math.Max(vw, vh) * 0.02;
            var nearby = all
                .Where(c => c.ContainsSvgPoint(svgX, svgY, tolerance))
                .OrderBy(c => (c.CenterX - svgX) * (c.CenterX - svgX) + (c.CenterY - svgY) * (c.CenterY - svgY))
                .FirstOrDefault();
            return (nearby, all);
        }
        catch (Exception ex) when (ex is PlaywrightException or HttpRequestException or TaskCanceledException or JsonException or XmlException)
        {
            _logger.LogDebug(ex, "[OnestopSeat] 사용자 클릭→후보+all 해석 실패.");
            return (null, null);
        }
    }

    // 수동 모드에서 구역 선택 대기 진입 시 호출한다. document 레벨 mousedown 리스너를
    // pre-install 하여 사용자가 blockImg 영역 내부를 클릭하는 순간의 좌표/이미지 rect/SVG URL을
    // 캡처한다. 구역 선택 완료 후 blockImg가 DOM에서 제거되어도 window.__nolUserClick이 유지되므로
    // 좌석 로드 감지 후 좌표→SVG 후보 매칭이 가능하다.
    private static async Task InstallUserZoneClickListenerAsync(IPage page)
    {
        try
        {
            await page.EvaluateAsync(@"() => {
                // 이전 핸들러 제거 (재진입 대응)
                if (window.__nolUserClickHandler) {
                    document.removeEventListener('mousedown', window.__nolUserClickHandler, true);
                }
                window.__nolUserClick = null;
                window.__nolUserClickHandler = (e) => {
                    const wrap = document.querySelector('[class*=""SeatMap_blockImg_""]');
                    if (!wrap) return;
                    const img = wrap.querySelector('img');
                    if (!img) return;
                    const rect = img.getBoundingClientRect();
                    if (rect.width <= 0 || rect.height <= 0) return;
                    if (e.clientX >= rect.left && e.clientX <= rect.right &&
                        e.clientY >= rect.top && e.clientY <= rect.bottom) {
                        window.__nolUserClick = {
                            x: e.clientX,
                            y: e.clientY,
                            imgLeft: rect.left,
                            imgTop: rect.top,
                            imgW: rect.width,
                            imgH: rect.height,
                            svgSrc: img.getAttribute('src') || '',
                            t: Date.now(),
                        };
                    }
                };
                document.addEventListener('mousedown', window.__nolUserClickHandler, true);
            }");
        }
        catch (PlaywrightException) { }
    }

    // 캡처된 사용자 클릭 좌표를 SVG 좌표계로 변환한 뒤, SVG 후보들의 bounding box 중 포함하는 것을 찾아
    // 해당 후보 자체(fill + bounds)를 반환한다. 반환된 candidate의 bounds를 사용해 좌석 선택에서
    // "같은 fill의 여러 구역" 중 정확한 구역만 필터할 수 있다.
    private async Task<NolZoneSvgCandidate?> ResolveUserClickZoneCandidateAsync(IPage page)
    {
        try
        {
            var clickJson = await page.EvaluateAsync<string?>(@"() => {
                const c = window.__nolUserClick;
                return c ? JSON.stringify(c) : null;
            }");
            if (string.IsNullOrEmpty(clickJson))
                return null;

            using var doc = JsonDocument.Parse(clickJson);
            var root = doc.RootElement;
            var clickX = root.GetProperty("x").GetDouble();
            var clickY = root.GetProperty("y").GetDouble();
            var imgLeft = root.GetProperty("imgLeft").GetDouble();
            var imgTop = root.GetProperty("imgTop").GetDouble();
            var imgW = root.GetProperty("imgW").GetDouble();
            var imgH = root.GetProperty("imgH").GetDouble();
            var svgSrc = root.GetProperty("svgSrc").GetString();

            if (imgW <= 0 || imgH <= 0 || string.IsNullOrEmpty(svgSrc))
                return null;

            var svgText = await _httpClient.GetStringAsync(svgSrc);
            var candidates = BuildNolOnestopZoneCandidatesFromSvg(svgText).ToList();
            if (candidates.Count == 0)
                return null;

            var vw = candidates[0].ViewWidth;
            var vh = candidates[0].ViewHeight;
            if (vw <= 0 || vh <= 0)
                return null;

            // 클릭 좌표를 SVG viewBox 좌표로 변환
            var svgX = (clickX - imgLeft) / imgW * vw;
            var svgY = (clickY - imgTop) / imgH * vh;

            // 매칭: bounds가 겹치는 경우(중첩 구역)가 많으므로 단순 area 우선은 오판한다.
            // 실측(초록 스탠딩 A idx10 bounds vs 파랑 idx14 bounds가 클릭점에 모두 걸림 → area 작은 파랑이 선택되는 버그).
            // 정확한 판정을 위해 (1) 포함하는 후보 중 center 거리가 가장 짧은 것을 우선한다.
            var containing = candidates
                .Where(c => c.ContainsSvgPoint(svgX, svgY))
                .OrderBy(c => (c.CenterX - svgX) * (c.CenterX - svgX) + (c.CenterY - svgY) * (c.CenterY - svgY))
                .ToList();
            if (containing.Count > 0)
            {
                var best = containing[0];
                if (containing.Count > 1)
                {
                    // 중첩된 경우 디버깅용 top-3 기록
                    var top = containing.Take(3).Select(c => $"{c.Fill}@({c.CenterX:F0},{c.CenterY:F0},area={c.Area:F0})").ToArray();
                    _logger.LogInformation("[OnestopSeat] 클릭({CX:F0},{CY:F0})→SVG({SX:F0},{SY:F0}) 중첩 {Count}개 후보. 최근접 선택: {Top}",
                        clickX, clickY, svgX, svgY, containing.Count, string.Join(" | ", top));
                }
                else
                {
                    _logger.LogDebug("[OnestopSeat] 클릭→SVG({SX:F1},{SY:F1}) 단독 매칭 fill={Fill}, key={Key}.", svgX, svgY, best.Fill, best.Key);
                }
                return best;
            }

            // 2차: 근처 후보 (오차 허용) — 영역 경계 바로 밖이면 viewBox 크기의 2% 허용
            var tolerance = Math.Max(vw, vh) * 0.02;
            var nearby = candidates
                .Where(c => c.ContainsSvgPoint(svgX, svgY, tolerance))
                .OrderBy(c => (c.CenterX - svgX) * (c.CenterX - svgX) + (c.CenterY - svgY) * (c.CenterY - svgY))
                .FirstOrDefault();
            if (nearby is not null)
            {
                _logger.LogDebug("[OnestopSeat] 클릭 좌표 근접(tol={Tol:F1}) 매칭 fill={Fill}, key={Key}.", tolerance, nearby.Fill, nearby.Key);
                return nearby;
            }

            _logger.LogDebug("[OnestopSeat] 클릭 좌표({CX:F1},{CY:F1})→SVG({SX:F1},{SY:F1})에 매칭되는 후보 없음 (후보 {Count}개).",
                clickX, clickY, svgX, svgY, candidates.Count);
            return null;
        }
        catch (Exception ex) when (ex is PlaywrightException or HttpRequestException or TaskCanceledException or JsonException or XmlException)
        {
            _logger.LogDebug(ex, "[OnestopSeat] 사용자 클릭 좌표→후보 해석 실패.");
            return null;
        }
    }

    // 좌석 클릭 후 NOL이 띄우는 alertdialog(ModalConfirm_outerWrap)을 감지해 "확인" 버튼을 자동 클릭한다.
    // 이 모달이 열려 있으면 '선택 완료' 버튼을 덮어 클릭이 계속 실패하므로 진입 즉시 닫아야 한다.
    private static async Task<bool> DismissNolSeatConfirmModalAsync(IPage page)
    {
        try
        {
            return await page.EvaluateAsync<bool>(@"() => {
                // ModalConfirm_outerWrap + role=alertdialog 조합을 우선 매칭
                let modal = document.querySelector('[class*=""ModalConfirm_outerWrap""][aria-modal=""true""], [class*=""ModalConfirm_""][role=""alertdialog""], div[role=""alertdialog""][aria-modal=""true""]');
                if (!modal) {
                    // 드문 변형: outerWrap 없이 내부 컨테이너만 있을 수 있음
                    const any = document.querySelector('[class*=""ModalConfirm_""]');
                    if (any && any.offsetWidth > 0) modal = any;
                }
                if (!modal) return false;

                // 확인 버튼 후보: primary 스타일 > '확인' 텍스트 > 첫 번째 button
                const primaries = [...modal.querySelectorAll('button[class*=""primary""], button[class*=""Primary""]')];
                const textConfirm = [...modal.querySelectorAll('button')].filter(b => {
                    const t = (b.innerText || '').trim();
                    return /^(\s*확인\s*|\s*예\s*|\s*OK\s*|\s*Confirm\s*)$/i.test(t);
                });
                const allButtons = [...modal.querySelectorAll('button')];

                const ordered = [...new Set([...primaries, ...textConfirm, ...allButtons])];
                for (const b of ordered) {
                    const s = window.getComputedStyle(b);
                    if (s.display === 'none' || s.visibility === 'hidden') continue;
                    if (b.disabled) continue;
                    b.click();
                    return true;
                }
                return false;
            }");
        }
        catch (PlaywrightException) { return false; }
    }

    // 활성 좌석 circle의 fill 분포를 분석해 지배적 fill(= 사용자가 선택한 구역 fill)을 판정한다.
    // 사용자 클릭 좌표 캡처가 실패한 극소수 케이스 fallback 용. 다중 구역 활성 케이스에서는
    // 좌석 수가 많은 구역이 답이 아닐 수 있어 정확도가 낮다(이전 버전의 주 버그 원인).
    // - max fill 개수가 2위보다 2배 이상이고 최소 20개 이상이어야 안정적 판정으로 간주.
    // - 판정 실패(경쟁 fill 존재 or 너무 적음)이면 null 반환 → 상위에서 fill 제한 없이 진행.
    private static async Task<string?> DetectDominantActiveSeatFillAsync(IPage page)
    {
        try
        {
            var json = await page.EvaluateAsync<string>(@"(() => {
                const circles = document.querySelectorAll('[class*=""SeatMap_seatGroup""] circle');
                const counts = {};
                for (const c of circles) {
                    const cls = c.getAttribute('class') || '';
                    if (cls.includes('disabled')) continue;
                    const r = parseFloat(c.getAttribute('r') || '0');
                    if (r <= 0) continue;
                    const cx = parseFloat(c.getAttribute('cx') || '0');
                    const cy = parseFloat(c.getAttribute('cy') || '0');
                    if (cx === 0 && cy === 0 && r <= 1) continue;
                    const s = window.getComputedStyle(c);
                    if (s.pointerEvents === 'none' || s.display === 'none' || s.visibility === 'hidden') continue;
                    const opacity = Number(s.opacity || '1');
                    if (opacity <= 0) continue;
                    const fill = (c.getAttribute('fill') || '').toLowerCase();
                    counts[fill] = (counts[fill] || 0) + 1;
                }
                return JSON.stringify(counts);
            })()");
            using var doc = JsonDocument.Parse(json);
            var entries = doc.RootElement.EnumerateObject()
                .Select(p => (Fill: p.Name, Count: p.Value.GetInt32()))
                .OrderByDescending(x => x.Count)
                .ToList();
            if (entries.Count == 0) return null;
            var top = entries[0];
            var second = entries.Count > 1 ? entries[1].Count : 0;
            if (top.Count >= 20 && top.Count >= second * 2)
                return top.Fill;
            return null;
        }
        catch (PlaywrightException) { return null; }
        catch (JsonException) { return null; }
    }

    // requiredFill 이 지정되면 해당 fill(소문자, 예: "#1ca814")에 일치하는 활성 circle만 센다.
    // 수동/자동 모드 모두에서 구역 선택 후 "선택된 구역" 좌석 수로 로드 완료를 판정하기 위해 사용한다.
    // 전체 활성 count가 필요하면 requiredFill=null 로 호출(=수동 모드 초기 대기).
    private static async Task<int> CountAvailableNolOnestopSeatsAsync(IPage page, string? requiredFill)
    {
        try
        {
            var fillArg = requiredFill?.ToLowerInvariant() ?? string.Empty;
            return await page.EvaluateAsync<int>(@"(fill) => {
                const circles = document.querySelectorAll('[class*=""SeatMap_seatGroup""] circle');
                let count = 0;
                for (const c of circles) {
                    const cls = c.getAttribute('class') || '';
                    if (cls.includes('disabled')) continue;
                    const r = parseFloat(c.getAttribute('r') || '0');
                    if (r <= 0) continue;
                    const cx = parseFloat(c.getAttribute('cx') || '0');
                    const cy = parseFloat(c.getAttribute('cy') || '0');
                    if (cx === 0 && cy === 0 && r <= 1) continue; // dummy helper circle 제외
                    const s = window.getComputedStyle(c);
                    if (s.pointerEvents === 'none' || s.display === 'none' || s.visibility === 'hidden') continue;
                    const opacity = Number(s.opacity || '1');
                    if (opacity <= 0) continue;
                    if (fill && (c.getAttribute('fill') || '').toLowerCase() !== fill) continue;
                    count++;
                }
                return count;
            }", fillArg);
        }
        catch (PlaywrightException) { return 0; }
    }

    private static async Task<(double width, double height)> GetNolViewportSizeAsync(IPage page)
    {
        try
        {
            var json = await page.EvaluateAsync<string>("() => JSON.stringify({w: window.innerWidth, h: window.innerHeight})");
            using var doc = JsonDocument.Parse(json);
            return (doc.RootElement.GetProperty("w").GetDouble(), doc.RootElement.GetProperty("h").GetDouble());
        }
        catch (PlaywrightException) { return (1920, 1080); }
    }

    private static bool IsViewportPoint(double x, double y, (double width, double height) viewport)
    {
        if (double.IsNaN(x) || double.IsNaN(y)) return false;
        if (x < 0 || y < 0) return false;
        if (x > viewport.width || y > viewport.height) return false;
        return true;
    }

    private static async Task<bool> WaitForNolOnestopSeatsLoadedAsync(IPage page, string? requiredFill, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (await CountAvailableNolOnestopSeatsAsync(page, requiredFill) >= 10)
                return true;
            // 30ms 폴링: 실측 구역 클릭→좌석 렌더 ~300ms. 30ms면 1~2 iteration 내 감지.
            await Task.Delay(30, cancellationToken);
        }
        return false;
    }

    // 화면 좌표(ClientX/Y, fill/area) + SVG 좌표(bounds) + 모든 후보 목록을 함께 반환.
    // 좌석 선택에서 'exclusive zone membership'(좌석을 포함하는 가장 작은 bbox가 selected zone인지) 판정에 필요.
    private async Task<(NolOnestopZoneCandidate? Zone, NolZoneSvgCandidate? Svg, IReadOnlyList<NolZoneSvgCandidate>? All)> TryFindNolOnestopZoneCandidateTripleAsync(IPage page, string? desiredBlock, HashSet<string> attemptedZones)
    {
        const string script = @"() => {
            const wrapper = document.querySelector('[class*=""SeatMap_blockImg_""]');
            const img = wrapper?.querySelector('img');
            if (!wrapper || !img || !img.src)
                return null;

            const rect = img.getBoundingClientRect();
            if (rect.width <= 0 || rect.height <= 0)
                return null;

            return {
                src: img.getAttribute('src') || '',
                left: rect.left,
                top: rect.top,
                width: rect.width,
                height: rect.height,
                naturalWidth: img.naturalWidth || 0,
                naturalHeight: img.naturalHeight || 0
            };
        }";

        try
        {
            var imageInfo = await page.EvaluateAsync<JsonElement?>(script);
            if (imageInfo is null || imageInfo.Value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
                return (null, null, null);

            var image = imageInfo.Value;
            var src = image.GetProperty("src").GetString();
            if (string.IsNullOrWhiteSpace(src))
                return (null, null, null);

            var left = image.GetProperty("left").GetDouble();
            var top = image.GetProperty("top").GetDouble();
            var width = image.GetProperty("width").GetDouble();
            var height = image.GetProperty("height").GetDouble();

            var svgText = await _httpClient.GetStringAsync(src);
            var allCandidates = BuildNolOnestopZoneCandidatesFromSvg(svgText).ToList();
            var ordered = allCandidates
                .OrderBy(x => x.CenterY)
                .ThenBy(x => x.CenterX)
                .ToList();

            if (ordered.Count == 0)
                return (null, null, null);

            // DesiredBlock이 "1","2"…처럼 1-based 정수 문자열이면 해당 후보를 최우선 시도.
            // 실측 기준 SVG에는 블록 이름(A1,B2 등) 메타데이터가 없어 텍스트 매칭은 불가능하다.
            var preferred = TryResolveNolZonePreferredIndex(desiredBlock, ordered.Count);

            foreach (var idx in EnumerateNolZoneTryOrder(preferred, ordered.Count))
            {
                var cand = ordered[idx];
                if (attemptedZones.Contains(cand.Key))
                    continue;

                var clientX = left + (cand.CenterX / cand.ViewWidth) * width;
                var clientY = top + (cand.CenterY / cand.ViewHeight) * height;
                var zone = new NolOnestopZoneCandidate(cand.Key, clientX, clientY, cand.Fill, cand.Area);
                return (zone, cand, allCandidates);
            }

            return (null, null, allCandidates);
        }
        catch (Exception ex) when (ex is PlaywrightException or HttpRequestException or TaskCanceledException or XmlException)
        {
            _logger.LogWarning(ex, "[OnestopSeat] 구역 후보 계산 실패");
            return (null, null, null);
        }
    }

    private static int? TryResolveNolZonePreferredIndex(string? desiredBlock, int total)
    {
        if (total <= 0 || string.IsNullOrWhiteSpace(desiredBlock))
            return null;

        var trimmed = desiredBlock.Trim();
        if (int.TryParse(trimmed, NumberStyles.None, CultureInfo.InvariantCulture, out var oneBased) &&
            oneBased >= 1 && oneBased <= total)
        {
            return oneBased - 1;
        }

        return null;
    }

    private static IEnumerable<int> EnumerateNolZoneTryOrder(int? preferredIndex, int total)
    {
        if (preferredIndex is int p && p >= 0 && p < total)
        {
            yield return p;
            for (var i = 0; i < total; i++)
            {
                if (i != p) yield return i;
            }
        }
        else
        {
            for (var i = 0; i < total; i++) yield return i;
        }
    }

    private sealed record NolZoneSvgCandidate(
        string Key,
        string Fill,
        double Area,
        double X,
        double Y,
        double Width,
        double Height,
        double ViewWidth,
        double ViewHeight)
    {
        public double CenterX => X + (Width / 2);
        public double CenterY => Y + (Height / 2);

        public bool ContainsSvgPoint(double px, double py, double tolerance = 0)
        {
            return px >= X - tolerance && px <= X + Width + tolerance
                && py >= Y - tolerance && py <= Y + Height + tolerance;
        }
    }

    private static IEnumerable<NolZoneSvgCandidate> BuildNolOnestopZoneCandidatesFromSvg(string svgText)
    {
        var document = XDocument.Parse(svgText);
        var root = document.Root;
        if (root is null)
            yield break;

        var (viewWidth, viewHeight) = ParseSvgViewBox(root);
        foreach (var element in root.Descendants())
        {
            var fill = (element.Attribute("fill")?.Value ?? string.Empty).Trim();
            if (!IsSelectableNolZoneFill(fill))
                continue;

            var bbox = element.Name.LocalName switch
            {
                "rect" => TryParseRectBounds(element),
                "polygon" => TryParsePolygonBounds(element),
                "path" => TryParsePathBounds(element),
                _ => null
            };

            if (bbox is null)
                continue;

            var (x, y, width, height) = bbox.Value;
            var area = width * height;
            if (area < 200)
                continue;

            var key = string.Create(CultureInfo.InvariantCulture, $"{fill}:{x:F2}:{y:F2}:{width:F2}:{height:F2}");
            yield return new NolZoneSvgCandidate(key, fill, area, x, y, width, height, viewWidth, viewHeight);
        }
    }

    private static bool IsSelectableNolZoneFill(string fill)
    {
        if (string.IsNullOrWhiteSpace(fill))
            return false;

        // 실측 기준(코다라인 26004589):
        //  - 선택 가능: #FB7E4E(주황 A/B), #17B3FF(파랑 C), #1CA814(초록 스탠딩), #7C68EE(보라 스탠딩 상단)
        //  - 선택 불가: 배경/경계선/텍스트 (white, black, #616161 등)
        // 단, 다른 공연 대응을 위해 블랙리스트 형태로 유지한다.
        return fill.Trim().ToLowerInvariant() switch
        {
            "#edeff3" or "#cacfda" or "#616161" or
            "#ffffff" or "#fff" or "white" or
            "#000000" or "#000" or "black" or
            "none" or "transparent" => false,
            _ => true
        };
    }

    private static (double ViewWidth, double ViewHeight) ParseSvgViewBox(XElement root)
    {
        var viewBox = root.Attribute("viewBox")?.Value;
        if (!string.IsNullOrWhiteSpace(viewBox))
        {
            var parts = viewBox.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 4 &&
                double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out var width) &&
                double.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out var height))
                return (width, height);
        }

        return (
            double.TryParse(root.Attribute("width")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var w) ? w : 421,
            double.TryParse(root.Attribute("height")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var h) ? h : 290);
    }

    private static (double X, double Y, double Width, double Height)? TryParseRectBounds(XElement element)
    {
        if (!double.TryParse(element.Attribute("x")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var x)) return null;
        if (!double.TryParse(element.Attribute("y")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var y)) y = 0;
        if (!double.TryParse(element.Attribute("width")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var width)) return null;
        if (!double.TryParse(element.Attribute("height")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var height)) return null;
        return (x, y, width, height);
    }

    private static (double X, double Y, double Width, double Height)? TryParsePolygonBounds(XElement element)
    {
        var points = element.Attribute("points")?.Value;
        if (string.IsNullOrWhiteSpace(points)) return null;
        var matches = NolSvgNumberPattern.Matches(points);
        if (matches.Count < 4) return null;

        var xs = new List<double>();
        var ys = new List<double>();
        for (var i = 0; i + 1 < matches.Count; i += 2)
        {
            xs.Add(double.Parse(matches[i].Value, CultureInfo.InvariantCulture));
            ys.Add(double.Parse(matches[i + 1].Value, CultureInfo.InvariantCulture));
        }

        return (xs.Min(), ys.Min(), xs.Max() - xs.Min(), ys.Max() - ys.Min());
    }

    private static (double X, double Y, double Width, double Height)? TryParsePathBounds(XElement element)
    {
        var d = element.Attribute("d")?.Value;
        if (string.IsNullOrWhiteSpace(d)) return null;
        var matches = NolSvgNumberPattern.Matches(d);
        if (matches.Count < 4) return null;

        var xs = new List<double>();
        var ys = new List<double>();
        for (var i = 0; i + 1 < matches.Count; i += 2)
        {
            xs.Add(double.Parse(matches[i].Value, CultureInfo.InvariantCulture));
            ys.Add(double.Parse(matches[i + 1].Value, CultureInfo.InvariantCulture));
        }

        return (xs.Min(), ys.Min(), xs.Max() - xs.Min(), ys.Max() - ys.Min());
    }

    private async Task<string> SelectNolOnestopSeatAsync(IPage page, HashSet<string> excludedSeats, NolZoneSvgCandidate? zoneCandidate, IReadOnlyList<NolZoneSvgCandidate>? allZones, ManualResetEventSlim? pauseGate, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        const int maxRetries = 10;
        // 필터 단계(중요 순):
        // 1) requiredFill: 선택된 구역의 fill과 일치하는 좌석만.
        // 2) hasBounds + bounds 내부: 선택 구역 SVG bbox 안의 좌석만.
        // 3) **exclusive zone membership**: 좌석 cx/cy를 포함하는 전체 후보 중 가장 작은 bbox가 selected zone인 경우에만.
        //    큰 구역(G/H) bbox 내에 작은 구역(A1/A2/A3) bbox가 포함되는 NEW URL 좌석맵에서 작은 구역 좌석이
        //    큰 구역 후보에 섞이는 버그(G 선택→A1 좌석 오클릭)를 차단.
        // zoneCandidate가 bounds를 모르는 fallback(dummy "<fill-only>")이면 (1)만 적용.
        var normalizedFill = zoneCandidate?.Fill.ToLowerInvariant() ?? string.Empty;
        var hasBounds = zoneCandidate is not null
            && !double.IsNaN(zoneCandidate.X) && !double.IsNaN(zoneCandidate.Y)
            && zoneCandidate.Width > 0 && zoneCandidate.Height > 0;
        var boundsX = hasBounds ? zoneCandidate!.X : double.NaN;
        var boundsY = hasBounds ? zoneCandidate!.Y : double.NaN;
        var boundsW = hasBounds ? zoneCandidate!.Width : double.NaN;
        var boundsH = hasBounds ? zoneCandidate!.Height : double.NaN;
        var selectedKey = hasBounds ? zoneCandidate!.Key : string.Empty;
        // exclusive membership용 모든 후보 배열 (같은 fill만 전달해 JS 계산량 축소)
        var zoneArray = hasBounds && allZones is not null
            ? allZones
                .Where(z => !string.IsNullOrEmpty(z.Fill) && z.Fill.Equals(zoneCandidate!.Fill, StringComparison.OrdinalIgnoreCase))
                .Where(z => !double.IsNaN(z.X) && !double.IsNaN(z.Y) && z.Width > 0 && z.Height > 0)
                .Select(z => new { key = z.Key, x = z.X, y = z.Y, w = z.Width, h = z.Height, area = z.Area })
                .ToArray()
            : Array.Empty<object>();
        for (var retry = 0; retry < maxRetries; retry++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                var excludedArray = excludedSeats.ToArray();
                var scanScript = @"(args) => {
                    const circles = document.querySelectorAll('__SEAT_SELECTOR__');
                    const excluded = args.excluded || [];
                    const requiredFill = (args.requiredFill || '').toLowerCase();
                    const hasBounds = !!args.hasBounds;
                    const bx = args.boundsX, by = args.boundsY, bw = args.boundsW, bh = args.boundsH;
                    const selectedKey = args.selectedKey || '';
                    const zones = args.zones || [];
                    const seats = [];
                    let i = 0;
                    for (const c of circles) {
                        const cls = c.getAttribute('class') || '';
                        if (cls.includes('disabled') || cls.includes('selected') || cls.includes('active')) { i++; continue; }
                        const style = window.getComputedStyle(c);
                        const pointerEvents = style.pointerEvents || '';
                        const opacity = Number(style.opacity || '1');
                        const visibility = style.visibility || '';
                        const display = style.display || '';
                        const radius = parseFloat(c.getAttribute('r') || '0');
                        if (pointerEvents === 'none' || opacity <= 0 || visibility === 'hidden' || display === 'none' || radius <= 0) { i++; continue; }
                        if (requiredFill && (c.getAttribute('fill') || '').toLowerCase() !== requiredFill) { i++; continue; }
                        const cx = parseFloat(c.getAttribute('cx') || '0');
                        const cy = parseFloat(c.getAttribute('cy') || '0');
                        if (hasBounds && (cx < bx || cx > bx + bw || cy < by || cy > by + bh)) { i++; continue; }
                        // exclusive zone membership: 좌석을 포함하는 전체 동색 구역 중 가장 작은 bbox가 selected와 같은지.
                        if (zones.length > 0 && selectedKey) {
                            let bestKey = null; let bestArea = Infinity;
                            for (const z of zones) {
                                if (cx >= z.x && cx <= z.x + z.w && cy >= z.y && cy <= z.y + z.h) {
                                    if (z.area < bestArea) { bestArea = z.area; bestKey = z.key; }
                                }
                            }
                            if (bestKey !== selectedKey) { i++; continue; }
                        }
                        const seatId = cx.toFixed(1) + ',' + cy.toFixed(1);
                        if (excluded.includes(seatId)) { i++; continue; }
                        seats.push({ cx: cx, cy: cy, idx: i, id: seatId });
                        i++;
                    }
                    if (seats.length === 0) return JSON.stringify({ s: 'empty', c: 0 });
                    seats.sort((a, b) => a.cy - b.cy || a.cx - b.cx);
                    const t = seats[0];
                    return JSON.stringify({ s: 'candidate', c: seats.length, id: t.id, idx: t.idx, cx: t.cx, cy: t.cy });
                }".Replace("__SEAT_SELECTOR__", NolOnestopSeatCircleSelector.Replace("'", "\\'"));
                var scanResult = await page.EvaluateAsync<string>(scanScript, new
                {
                    excluded = excludedArray,
                    requiredFill = normalizedFill,
                    hasBounds,
                    boundsX,
                    boundsY,
                    boundsW,
                    boundsH,
                    selectedKey,
                    zones = zoneArray,
                });

                using var doc = JsonDocument.Parse(scanResult);
                var status = doc.RootElement.GetProperty("s").GetString();
                var count = doc.RootElement.GetProperty("c").GetInt32();

                if (status == "empty")
                {
                    if (excludedSeats.Count > 0)
                    {
                        _logger.LogWarning("[OnestopSeat] 제외 좌석 빼면 선택 가능 좌석 없음 — 초기화. retry={Retry}", retry);
                        excludedSeats.Clear();
                    }
                    else
                        _logger.LogWarning("[OnestopSeat] 선택 가능 좌석 없음. retry={Retry}", retry);
                    await Task.Delay(50, cancellationToken);
                    continue;
                }

                if (status != "candidate")
                {
                    _logger.LogWarning("[OnestopSeat] 좌석 클릭 실패: status={Status}, count={Count}", status, count);
                    continue;
                }

                var seatId = doc.RootElement.GetProperty("id").GetString()!;
                var seatIndex = doc.RootElement.GetProperty("idx").GetInt32();
                var circleLocator = page.Locator(NolOnestopSeatCircleSelector).Nth(seatIndex);
                try { await circleLocator.ScrollIntoViewIfNeededAsync(); } catch (PlaywrightException) { }

                // zoom 확대 상태에서 좌석이 viewport 밖이면 scrollIntoView(center)로 viewport 중심에 가져온다.
                // EntZoomableWrapper 같은 transform container에서는 일반 scrollTo가 안 통해도
                // element.scrollIntoView({block:'center'})는 브라우저가 적절히 처리.
                try
                {
                    await page.EvaluateAsync(@"(idx) => {
                        const circles = document.querySelectorAll('__SEAT_SELECTOR__');
                        const c = circles[idx];
                        if (!c) return;
                        const r = c.getBoundingClientRect();
                        if (r.left < 0 || r.top < 0 || r.right > window.innerWidth || r.bottom > window.innerHeight) {
                            c.scrollIntoView({ block: 'center', inline: 'center' });
                        }
                    }".Replace("__SEAT_SELECTOR__", NolOnestopSeatCircleSelector.Replace("'", "\\'")), seatIndex);
                }
                catch (PlaywrightException) { }

                // 3단계 클릭 전략: Locator.ClickAsync(Force) → Mouse fallback → JS dispatchEvent
                var clicked = false;
                try
                {
                    await circleLocator.ClickAsync(new LocatorClickOptions { Force = true, Timeout = 1500 });
                    _logger.LogInformation("[OnestopSeat] 좌석 클릭 완료(Locator). available={Count}, seatId={SeatId}, idx={Index}", count, seatId, seatIndex);
                    clicked = true;
                }
                catch (PlaywrightException ex)
                {
                    _logger.LogDebug(ex, "[OnestopSeat] Locator.ClickAsync 실패 — Mouse fallback 시도.");
                }

                if (!clicked)
                {
                    var box = await circleLocator.BoundingBoxAsync();
                    if (box is not null && box.Width > 0 && box.Height > 0)
                    {
                        var seatX = (float)(box.X + (box.Width / 2));
                        var seatY = (float)(box.Y + (box.Height / 2));
                        if (seatX >= 0 && seatY >= 0)
                        {
                            try
                            {
                                await page.Mouse.MoveAsync(seatX, seatY);
                                await page.Mouse.DownAsync();
                                await page.Mouse.UpAsync();
                                _logger.LogInformation("[OnestopSeat] 좌석 클릭 완료(Mouse fallback). available={Count}, seatId={SeatId}, idx={Index}, pt=({X:F1},{Y:F1})", count, seatId, seatIndex, seatX, seatY);
                                clicked = true;
                            }
                            catch (PlaywrightException) { }
                        }
                    }
                }

                if (!clicked)
                {
                    // 최후 수단: JS element.dispatchEvent로 synthetic MouseEvent 발사. trusted=false지만
                    // React synthetic handler는 일반적으로 처리. viewport 밖에서도 동작.
                    try
                    {
                        var jsClicked = await page.EvaluateAsync<bool>(@"(idx) => {
                            const circles = document.querySelectorAll('__SEAT_SELECTOR__');
                            const c = circles[idx];
                            if (!c) return false;
                            const r = c.getBoundingClientRect();
                            const x = r.left + r.width / 2;
                            const y = r.top + r.height / 2;
                            const ev = new MouseEvent('click', { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, button: 0 });
                            c.dispatchEvent(ev);
                            return true;
                        }".Replace("__SEAT_SELECTOR__", NolOnestopSeatCircleSelector.Replace("'", "\\'")), seatIndex);
                        if (jsClicked)
                        {
                            _logger.LogInformation("[OnestopSeat] 좌석 클릭 완료(JS dispatchEvent). available={Count}, seatId={SeatId}, idx={Index}", count, seatId, seatIndex);
                            clicked = true;
                        }
                    }
                    catch (PlaywrightException ex)
                    {
                        _logger.LogDebug(ex, "[OnestopSeat] JS dispatchEvent 실패.");
                    }
                }

                if (!clicked)
                {
                    excludedSeats.Add(seatId);
                    _logger.LogWarning("[OnestopSeat] 좌석 클릭 3단계 전부 실패 — 제외 후 다음 좌석. seatId={SeatId}", seatId);
                    continue;
                }
                // 좌석 클릭 후 DOM 상태 업데이트 최소 대기. 너무 짧으면 완료 버튼 아직 active 안 됨.
                await Task.Delay(30, cancellationToken);

                // NOL은 좌석 클릭 시 종종 "선택하시겠습니까?" 같은 ModalConfirm_outerWrap 확인 모달을
                // 띄운다. 이 모달이 '선택 완료' 버튼 위를 덮어 pointer events를 가로채므로
                // 완료 클릭이 계속 timeout 1500ms * 재시도로 수초씩 밀리는 지연의 주범이 된다.
                // 감지 즉시 확인 버튼을 자동 클릭해 모달을 닫는다.
                var dismissed = await DismissNolSeatConfirmModalAsync(page);
                if (dismissed)
                {
                    _logger.LogInformation("[OnestopSeat] 좌석 선택 확인 모달 자동 확인 클릭.");
                    // 모달 close 애니메이션/DOM detach 완료 대기
                    await Task.Delay(80, cancellationToken);
                }

                // selectionConfirmed 500ms: React state 반영 시간. 성공은 보통 100~200ms, 500ms면 2배 여유.
                var selectionConfirmed = await PlaywrightRuntime.TryWaitForConditionAsync(
                    async () =>
                    {
                        try
                        {
                            var infoLocator = page.Locator("[class*='InfoSelected'], [class*='infoSelected']").First;
                            if (await infoLocator.CountAsync() > 0)
                            {
                                var text = await infoLocator.InnerTextAsync();
                                if (!text.Contains("선택한 좌석이 없습니다", StringComparison.OrdinalIgnoreCase))
                                    return true;
                            }

                            var completeBtn = page.Locator("button:has-text('선택 완료'), [class*='EntButton'] button:has-text('완료')").First;
                            if (await completeBtn.CountAsync() > 0 && !await completeBtn.IsDisabledAsync())
                                return true;

                            var className = await circleLocator.GetAttributeAsync("class") ?? string.Empty;
                            return className.Contains("selected", StringComparison.OrdinalIgnoreCase) ||
                                   className.Contains("active", StringComparison.OrdinalIgnoreCase);
                        }
                        catch { return false; }
                    },
                    TimeSpan.FromMilliseconds(500), cancellationToken);

                if (!selectionConfirmed)
                {
                    excludedSeats.Add(seatId);
                    _logger.LogWarning("[OnestopSeat] 좌석 선택 미반영 (중복 추정). seatId={SeatId}, retry={Retry}, excludedCount={ExcludedCount}",
                        seatId, retry, excludedSeats.Count);

                    if (pauseGate is not null)
                    {
                        pauseGate.Reset();
                        _logger.LogInformation("[NOL] 중복 좌석 감지 후 일시정지 — 사용자 재개 대기 중. excluded={SeatId}", seatId);
                        progress?.Report(new AutomationProgress("중복 감지 — 일시정지", $"중복 좌석 감지됨 ({seatId}). 재개 버튼을 눌러주세요."));
                        await Task.Run(() => pauseGate.Wait(cancellationToken), cancellationToken);
                        _logger.LogInformation("[NOL] 중복 감지 일시정지 해제 — 다른 좌석 선택 진행.");
                        progress?.Report(new AutomationProgress("좌석 재선택 중", "일시정지 해제 — 다른 좌석 선택 진행"));
                    }

                    continue;
                }

                return seatId;
            }
            catch (PlaywrightException ex)
            {
                _logger.LogWarning("[OnestopSeat] 좌석 선택 중 예외. retry={Retry}, error={Error}", retry, ex.Message);
                await Task.Delay(50, cancellationToken);
            }
        }
        throw new InvalidOperationException($"원스탑 좌석 선택 실패 ({maxRetries}회 시도).");
    }

    private async Task ClickNolOnestopSeatCompleteAsync(IPage page, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var completeBtn = page.Locator("button:has-text('선택 완료'), [class*='EntButton'] button:has-text('완료')").First;

        await PlaywrightRuntime.WaitForConditionAsync(
            async () =>
            {
                try
                {
                    if (await completeBtn.CountAsync() == 0) return false;
                    return !await completeBtn.IsDisabledAsync();
                }
                catch (PlaywrightException) { return false; }
            },
            timeout, cancellationToken, "NOL '선택 완료' 버튼이 활성화되지 않았습니다.");

        try { await completeBtn.ScrollIntoViewIfNeededAsync(); }
        catch (PlaywrightException) { }

        var beforeUrl = page.Url;

        // Playwright ClickAsync는 CDP Input.dispatchMouseEvent 기반 trusted click을 생성한다.
        // Timeout 1500ms: 일반 버튼 클릭은 100~300ms 내 완료. 1.5초면 로딩 상태도 대응.
        try
        {
            await completeBtn.ClickAsync(new LocatorClickOptions { Force = true, Timeout = 1500 });
        }
        catch (PlaywrightException ex)
        {
            _logger.LogWarning(ex, "[OnestopSeat] '선택 완료' ClickAsync(Force) 실패 — ClickAsync(일반) 재시도.");
            await completeBtn.ClickAsync(new LocatorClickOptions { Timeout = 1500 });
        }
        _logger.LogInformation("[OnestopSeat] '선택 완료' 버튼 Playwright trusted click 수행.");

        if (await IsOnestopSeatCompleteConfirmedAsync(page, completeBtn, beforeUrl, cancellationToken))
        {
            _logger.LogInformation("[OnestopSeat] '선택 완료' 클릭 확인됨 (페이지 전환 또는 좌석 페이지 이탈).");
            return;
        }

        _logger.LogWarning("[OnestopSeat] '선택 완료' 첫 클릭 미반영 — 재시도.");
        try
        {
            await completeBtn.ClickAsync(new LocatorClickOptions { Timeout = 1500 });
        }
        catch (PlaywrightException)
        {
            await completeBtn.ClickAsync(new LocatorClickOptions { Force = true, Timeout = 1500 });
        }

        if (await IsOnestopSeatCompleteConfirmedAsync(page, completeBtn, beforeUrl, cancellationToken))
        {
            _logger.LogInformation("[OnestopSeat] '선택 완료' 재시도 클릭 확인됨.");
            return;
        }

        throw new InvalidOperationException("NOL '선택 완료' 버튼 클릭 후 페이지 전환이 확인되지 않았습니다. 좌석 선택은 완료되었으나 '선택 완료' 처리가 실패했습니다.");
    }

    private static async Task<bool> IsOnestopSeatCompleteConfirmedAsync(IPage page, ILocator completeBtn, string beforeUrl, CancellationToken cancellationToken)
    {
        return await PlaywrightRuntime.TryWaitForConditionAsync(
            async () =>
            {
                if (page.IsClosed) return true;

                var currentUrl = page.Url;
                if (!string.Equals(currentUrl, beforeUrl, StringComparison.OrdinalIgnoreCase) &&
                    !currentUrl.Contains("/onestop/seat", StringComparison.OrdinalIgnoreCase))
                    return true;

                try
                {
                    if (await page.Locator(NolOnestopSeatMapSelector).CountAsync() == 0 &&
                        await page.Locator(NolOnestopSeatCircleSelector).CountAsync() == 0)
                        return true;

                    if (await page.Locator("[class*='BookingProgress'], [class*='bookingProgress'], [class*='Payment'], [class*='payment'], [class*='Delivery'], [class*='delivery']").CountAsync() > 0)
                        return true;

                    if (await completeBtn.CountAsync() == 0)
                        return true;
                }
                catch (PlaywrightException ex) when (PlaywrightRuntime.IsClosedTargetError(ex))
                {
                    return true;
                }

                return false;
            },
            // 3500ms: 실측상 정상 전환은 300~500ms지만 서버/네트워크 variance가 커서
            // 2500ms로는 timeout 후 재시도가 자주 발생(실측 2회차 2886ms). 재시도 1회 오버헤드(~3000ms)가
            // timeout 1000ms 추가보다 훨씬 크므로 3500ms로 상향해 재시도를 방지.
            TimeSpan.FromMilliseconds(3500), cancellationToken);
    }

    private async Task SelectNolLegacySeatAndCompleteAsync(IPage page, TimeSpan timeout, IProgress<AutomationProgress>? progress, string? desiredBlock, HashSet<string> excludedSeats, ManualResetEventSlim? pauseGate, CancellationToken cancellationToken)
    {
        var stepSw = Stopwatch.StartNew();

        var seatFrame = await FindNolLegacyFrameAsync(page, "ifrmSeat", timeout, cancellationToken)
            ?? throw new TimeoutException("NOL 레거시 프레임(ifrmSeat)을 찾지 못했습니다.");
        _logger.LogInformation("[LegacySeat] ifrmSeat 프레임 발견. frameUrl={Url}, findMs={Ms}", seatFrame.Url, stepSw.ElapsedMilliseconds);

        stepSw.Restart();
        var detailFrame = await FindNolLegacyFrameAsync(page, "ifrmSeatDetail", TimeSpan.FromSeconds(3), cancellationToken, throwOnTimeout: false);
        _logger.LogInformation("[LegacySeat] ifrmSeatDetail 프레임={Found}. findMs={Ms}", detailFrame is not null, stepSw.ElapsedMilliseconds);

        var workFrame = detailFrame ?? seatFrame;

        stepSw.Restart();
        var hasSelectSeat = await HasNolLegacySelectSeatSpansAsync(workFrame);
        _logger.LogInformation("[LegacySeat] SelectSeat span 존재={Exists}. checkMs={Ms}", hasSelectSeat, stepSw.ElapsedMilliseconds);

        if (!hasSelectSeat)
        {
            stepSw.Restart();
            var modeLabel = string.IsNullOrWhiteSpace(desiredBlock) ? "사용자 수동 대기" : "자동 선택";
            progress?.Report(new AutomationProgress("구역 선택 중", $"구역 {modeLabel} 중"));
            await EnsureNolLegacyZoneSelectedAsync(workFrame, desiredBlock, progress, cancellationToken);
            _logger.LogInformation("[LegacySeat] 구역 선택 후 좌석 로드 완료. mode={Mode}, waitMs={Ms}", modeLabel, stepSw.ElapsedMilliseconds);
        }

        stepSw.Restart();
        progress?.Report(new AutomationProgress("좌석 선택 중"));
        await SelectNolLegacySeatAsync(workFrame, excludedSeats, pauseGate, progress, cancellationToken);
        _logger.LogInformation("[LegacySeat] 좌석 선택 완료. selectMs={Ms}", stepSw.ElapsedMilliseconds);

        stepSw.Restart();
        await ClickNolLegacySeatCompleteAsync(seatFrame, cancellationToken);
        _logger.LogInformation("[LegacySeat] 선택 완료 처리. completeMs={Ms}", stepSw.ElapsedMilliseconds);
    }

    private static async Task<IFrame?> FindNolLegacyFrameAsync(IPage page, string frameName, TimeSpan timeout, CancellationToken cancellationToken, bool throwOnTimeout = true)
    {
        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (var frame in page.Frames)
            {
                if (frame == page.MainFrame) continue;
                if (string.Equals(frame.Name, frameName, StringComparison.OrdinalIgnoreCase))
                    return frame;
            }
            await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
        }
        if (throwOnTimeout)
            throw new TimeoutException($"NOL 레거시 프레임({frameName})을 찾지 못했습니다.");
        return null;
    }

    private static async Task<bool> HasNolLegacySelectSeatSpansAsync(IFrame frame)
    {
        try
        {
            var count = await frame.EvaluateAsync<int>(@"() => {
                return document.querySelectorAll('span[onclick*=""SelectSeat""]').length;
            }");
            return count > 0;
        }
        catch (PlaywrightException) { return false; }
    }

    private async Task EnsureNolLegacyZoneSelectedAsync(IFrame frame, string? desiredBlock, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        // DesiredBlock이 비어있으면 사용자가 직접 구역을 클릭할 때까지 대기 (자동 클릭 금지).
        // 값이 있으면 기존처럼 자동 선택(첫 매칭 후보 클릭).
        if (string.IsNullOrWhiteSpace(desiredBlock))
        {
            _logger.LogInformation("[LegacySeat] DesiredBlock 비어있음 — 구역 수동 선택 대기 모드.");
            progress?.Report(new AutomationProgress("구역 선택 대기 중", "구역을 직접 클릭해 주세요 (DesiredBlock 미지정)."));

            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    var count = await frame.EvaluateAsync<int>(@"() => {
                        return document.querySelectorAll('span[onclick*=""SelectSeat""]').length;
                    }");
                    if (count > 0)
                    {
                        _logger.LogInformation("[LegacySeat] 사용자 구역 선택 감지 — SelectSeat span {Count}개 로드.", count);
                        return;
                    }
                }
                catch (PlaywrightException) { }
                // 25ms 폴링: 사용자 구역 클릭 즉시 감지. NEW 경로와 동일 기준.
                await Task.Delay(25, cancellationToken);
            }
        }

        _logger.LogInformation("[LegacySeat] 구역 자동 선택 시작. desiredBlock={Desired}", desiredBlock);
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                var count = await frame.EvaluateAsync<int>(@"() => {
                    return document.querySelectorAll('span[onclick*=""SelectSeat""]').length;
                }");
                if (count > 0)
                {
                    _logger.LogInformation("[LegacySeat] SelectSeat span {Count}개 감지.", count);
                    return;
                }

                if (await TrySelectNolLegacyZoneAsync(frame))
                {
                    _logger.LogInformation("[LegacySeat] 구역 후보 클릭 완료 — 좌석 로드 대기.");
                    // 구역 클릭 후 서버 응답(SelectSeat span 렌더)까지 시간 필요. 80ms 최소 대기.
                    await Task.Delay(80, cancellationToken);
                    continue;
                }
            }
            catch (PlaywrightException) { }
            await Task.Delay(25, cancellationToken);
        }
    }

    private static async Task<bool> TrySelectNolLegacyZoneAsync(IFrame frame)
    {
        const string script = @"() => {
            const selectors = [
                'area[href*=""Block""]',
                'area[href*=""block""]',
                'span[onclick*=""Block""]',
                'span[onclick*=""block""]',
                'a[href*=""Block""]',
                'a[href*=""block""]',
                'a[onclick*=""Block""]',
                'a[onclick*=""block""]'
            ];

            for (const selector of selectors) {
                const nodes = Array.from(document.querySelectorAll(selector));
                for (const node of nodes) {
                    const style = window.getComputedStyle(node);
                    if (style.display === 'none' || style.visibility === 'hidden' || style.pointerEvents === 'none') continue;
                    if (typeof node.click === 'function') {
                        node.click();
                        return true;
                    }
                    node.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
                    return true;
                }
            }

            return false;
        }";

        try
        {
            return await frame.EvaluateAsync<bool>(script);
        }
        catch (PlaywrightException)
        {
            return false;
        }
    }

    private async Task SelectNolLegacySeatAsync(IFrame workFrame, HashSet<string> excludedSeats, ManualResetEventSlim? pauseGate, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        const int maxRetries = 10;
        for (var retry = 0; retry < maxRetries; retry++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                var excludedArray = excludedSeats.ToArray();
                var scanResult = await workFrame.EvaluateAsync<string>(@"(args) => {
                    const excluded = args.excluded || [];
                    const spans = document.querySelectorAll('span[onclick*=""SelectSeat""]');
                    if (spans.length === 0) return JSON.stringify({ s: 'no_seats', c: 0 });
                    for (const span of spans) {
                        const onclick = span.getAttribute('onclick') || '';
                        const match = onclick.match(/SelectSeat\(this,'([^']*)','([^']*)','([^']*)','([^']*)'/);
                        const seatId = match ? match[3] + '_' + match[4] : onclick.substring(0, 50);
                        if (excluded.includes(seatId)) continue;
                        const style = window.getComputedStyle(span);
                        if (style.display === 'none' || style.visibility === 'hidden') continue;
                        span.click();
                        return JSON.stringify({ s: 'clicked', c: spans.length, id: seatId });
                    }
                    return JSON.stringify({ s: 'empty', c: spans.length });
                }", new { excluded = excludedArray });

                using var doc = JsonDocument.Parse(scanResult);
                var status = doc.RootElement.GetProperty("s").GetString();
                var count = doc.RootElement.GetProperty("c").GetInt32();

                if (status == "empty" || status == "no_seats")
                {
                    var hadExcluded = excludedSeats.Count > 0;
                    if (hadExcluded)
                    {
                        _logger.LogWarning("[LegacySeat] 제외 좌석 {ExcludedCount}개 빼면 선택 가능 좌석 없음 — 초기화. retry={Retry}", excludedSeats.Count, retry);
                        excludedSeats.Clear();

                        if (pauseGate is not null)
                        {
                            pauseGate.Reset();
                            _logger.LogInformation("[NOL] 모든 시도 좌석 중복 — 일시정지 — 사용자 재개 대기 중.");
                            progress?.Report(new AutomationProgress("중복 감지 — 일시정지", "시도한 좌석 모두 중복. 재개 버튼을 눌러주세요."));
                            await Task.Run(() => pauseGate.Wait(cancellationToken), cancellationToken);
                            _logger.LogInformation("[NOL] 일시정지 해제 — 좌석 재선택 진행.");
                            progress?.Report(new AutomationProgress("좌석 재선택 중", "일시정지 해제 — 좌석 재선택 진행"));
                        }
                    }
                    else
                    {
                        _logger.LogWarning("[LegacySeat] 선택 가능 좌석 없음. count={Count}, retry={Retry}", count, retry);
                    }
                    await Task.Delay(100, cancellationToken);
                    continue;
                }

                if (status == "clicked")
                {
                    var seatId = doc.RootElement.GetProperty("id").GetString()!;
                    // 좌석 클릭 후 DOM 업데이트 최소 대기. 50ms면 상태 반영 충분.
                    await Task.Delay(50, cancellationToken);

                    // 좌석 선택 확인: poticket DOM은 공연/플레이스별로 selector 구조가 다르므로
                    // 포괄적 selector 집합 + SelectSeat span 자체의 상태 클래스까지 종합 검사한다.
                    var selectionConfirmed = await PlaywrightRuntime.TryWaitForConditionAsync(
                        async () =>
                        {
                            try
                            {
                                return await workFrame.EvaluateAsync<bool>(@"() => {
                                    // 1) 선택된 좌석 영역 (다양한 selector 커버)
                                    const listSel = [
                                        '#SelectedSeat li',
                                        '#divSelectedSeat li',
                                        '#divSelSeat li',
                                        '#divSelSeatDetail li',
                                        '[id*=""SelectedSeat""] li',
                                        '[id*=""selSeat"" i] li',
                                        '.selected_seat li'
                                    ].join(',');
                                    const listCount = document.querySelectorAll(listSel).length;
                                    if (listCount > 0) return true;

                                    // 2) SelectSeat span 자체의 선택 상태 클래스
                                    const stateSel = 'span[onclick*=""SelectSeat""].on, span[onclick*=""SelectSeat""].selected, span[onclick*=""SelectSeat""].active, span[onclick*=""SelectSeat""][class*=""selected""]';
                                    if (document.querySelectorAll(stateSel).length > 0) return true;

                                    // 3) 선택 개수 표시 input/label 등
                                    const cntEl = document.querySelector('#iSeatCnt, #txtSeatCnt, [id*=""SeatCnt""]');
                                    if (cntEl) {
                                        const v = parseInt((cntEl.value || cntEl.innerText || '0').replace(/\D/g, ''), 10);
                                        if (!isNaN(v) && v > 0) return true;
                                    }
                                    return false;
                                }");
                            }
                            catch { return false; }
                        },
                        // 400ms: 일반 poticket 좌석 선택 DOM 반영 ~200ms. 400ms면 여유.
                        TimeSpan.FromMilliseconds(400), cancellationToken);

                    // selector가 실제 DOM과 다를 수 있으므로 미확인 상태여도 완료 단계로 진행한다.
                    // 진짜 실패였다면 ClickNolLegacySeatCompleteAsync가 페이지 이동을 감지 못해 예외를 던지고
                    // 상위 재시도 루프에서 다른 좌석을 시도한다. 1회 실행당 좌석은 반드시 1개만 클릭한다.
                    if (selectionConfirmed)
                        _logger.LogInformation("[LegacySeat] 좌석 클릭+선택 확인 완료. count={Count}, seatId={SeatId}", count, seatId);
                    else
                        _logger.LogInformation("[LegacySeat] 좌석 클릭 완료(선택 확인 미매칭) — 완료 단계로 진행. seatId={SeatId}", seatId);

                    return;
                }
            }
            catch (PlaywrightException ex)
            {
                _logger.LogWarning("[LegacySeat] 좌석 선택 중 예외. retry={Retry}, error={Error}", retry, ex.Message);
                await Task.Delay(100, cancellationToken);
            }
        }
        throw new InvalidOperationException($"레거시 좌석 선택 실패 ({maxRetries}회 시도).");
    }

    private async Task ClickNolLegacySeatCompleteAsync(IFrame seatFrame, CancellationToken cancellationToken)
    {
        await Task.Delay(100, cancellationToken);

        var startUrl = seatFrame.Url;
        var clickedMethod = string.Empty;

        try
        {
            var result = await seatFrame.EvaluateAsync<bool>(@"() => {
                if (typeof fnSelect === 'function') { fnSelect(); return true; }
                return false;
            }");
            if (result)
                clickedMethod = "fnSelect";
        }
        catch (PlaywrightException) { }

        if (string.IsNullOrEmpty(clickedMethod))
        {
            try
            {
                var clicked = await seatFrame.EvaluateAsync<bool>(@"() => {
                    var img = document.querySelector('#NextStepImage');
                    if (img && img.parentElement) { img.parentElement.click(); return true; }
                    var links = document.querySelectorAll('a[href*=""fnSelect""]');
                    for (var l of links) { l.click(); return true; }
                    var btns = document.querySelectorAll('a[onclick*=""fnSelect""], button[onclick*=""fnSelect""], input[onclick*=""fnSelect""]');
                    for (var b of btns) { b.click(); return true; }
                    return false;
                }");
                if (clicked)
                    clickedMethod = "NextStepImage/link";
            }
            catch (PlaywrightException) { }
        }

        if (string.IsNullOrEmpty(clickedMethod))
            throw new InvalidOperationException("레거시 '좌석 선택 완료' 트리거를 찾지 못함 (fnSelect/NextStepImage/링크 모두 실패).");

        // 좌석 선택 완료가 실제로 처리되었는지 페이지(iframe) 전환으로 확인.
        // 구체적으로: ifrmSeat URL이 loading.html 또는 다른 단계로 바뀌거나,
        // 부모 페이지의 다른 단계 iframe(ifrmBookCertify/ifrmBookEnd 등)이 유효한 URL로 바뀜.
        // 1800ms: iframe 전환 일반 1초 내, 1.8초면 서버 지연 포함해도 충분.
        var transitioned = await WaitForNolLegacyCompleteTransitionAsync(seatFrame, startUrl, TimeSpan.FromMilliseconds(1800), cancellationToken);
        if (!transitioned)
        {
            throw new InvalidOperationException($"레거시 '좌석 선택 완료' 클릭({clickedMethod}) 후 페이지 전환이 감지되지 않음 — 좌석 선택이 반영되지 않았을 가능성이 높음.");
        }

        _logger.LogInformation("[LegacySeat] 좌석 선택 완료({Method}) — 페이지 전환 확인.", clickedMethod);
    }

    private static async Task<bool> WaitForNolLegacyCompleteTransitionAsync(IFrame seatFrame, string startUrl, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                var currentUrl = seatFrame.Url;
                if (!string.Equals(currentUrl, startUrl, StringComparison.OrdinalIgnoreCase))
                    return true;

                // seatFrame URL은 그대로여도 다른 단계 iframe들이 유효한 URL로 전환됐는지 확인
                var page = seatFrame.Page;
                foreach (var f in page.Frames)
                {
                    if (f == seatFrame || f == page.MainFrame) continue;
                    var url = f.Url ?? string.Empty;
                    // BookCertify.asp, BookEnd.asp, Payment 등 다음 단계 URL 등장 시 전환으로 간주
                    if (url.Contains("BookCertify", StringComparison.OrdinalIgnoreCase) &&
                        !url.Contains("loading", StringComparison.OrdinalIgnoreCase))
                        return true;
                    if (url.Contains("BookEnd", StringComparison.OrdinalIgnoreCase) &&
                        !url.Contains("loading", StringComparison.OrdinalIgnoreCase))
                        return true;
                    if (url.Contains("Payment", StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch (PlaywrightException) { }

            // 30ms 폴링: iframe URL/프레임 변화 빠르게 감지.
            await Task.Delay(30, cancellationToken);
        }
        return false;
    }

    private static string StripUrlFragment(string url)
    {
        var hashIndex = url.IndexOf('#');
        return hashIndex >= 0 ? url[..hashIndex] : url;
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
