using Microsoft.Extensions.Logging;
using Microsoft.Playwright;
using Polly;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using KillRiceMonkey.Application.Abstractions;
using KillRiceMonkey.Application.Models;

namespace KillRiceMonkey.Infrastructure.Services;

public sealed class Yes24AutomationService : IYes24AutomationService, IAsyncDisposable
{
    private const string Yes24RemoteDebugLaunchUrl = "https://ticket.yes24.com/";
    private const int Yes24RemoteDebugPort = 9224;
    private const string Yes24CdpEndpoint = "http://127.0.0.1:9224/";
    private const string Yes24ProfileName = "Yes24RemoteDebugProfile";
    private static readonly Regex ChangeBlockRegex = new(@"ChangeBlock\((\d+)\)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex PerfIdRegex = new(@"/Perf/(?<id>\d+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly string[] BrowserExecutableCandidates =
    [
        @"C:\Program Files\Google\Chrome\Application\chrome.exe",
        @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        @"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    ];

    private readonly ILogger<Yes24AutomationService> _logger;
    private readonly PlaywrightRuntime _runtime;
    private readonly ResiliencePipeline _pipeline;
    private readonly SemaphoreSlim _yes24BrowserLock = new(1, 1);
    private IBrowser? _preparedYes24ConnectedBrowser;
    private IPage? _preparedYes24Page;
    private IPage? _preparedYes24SalePopup;

    public Yes24AutomationService(
        ILogger<Yes24AutomationService> logger,
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
                return await RunYes24AutomationAsync(request, progress, token);
            }, cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "YES24 automation failed with exception");
            return new AutomationRunResult(false, $"예외 발생: {ex.Message}", DateTimeOffset.Now);
        }
    }

    public async ValueTask DisposeAsync()
    {
        await ReleasePreparedYes24ConnectionAsync();
        _yes24BrowserLock.Dispose();
    }

    public async Task<bool> IsRemoteDebugBrowserAvailableAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return await PlaywrightRuntime.IsCdpEndpointAvailableAsync(Yes24CdpEndpoint, cancellationToken);
    }

    public async Task<bool> IsAutomationPreparedAsync(CancellationToken cancellationToken)
    {
        return await TryGetPreparedYes24PageAsync(cancellationToken) is not null;
    }

    public async Task<string> LaunchRemoteDebugBrowserAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (await PlaywrightRuntime.IsCdpEndpointAvailableAsync(Yes24CdpEndpoint, cancellationToken))
        {
            return $"이미 remote debug 브라우저가 열려 있습니다: {Yes24CdpEndpoint}";
        }

        var executablePath = BrowserExecutableCandidates.FirstOrDefault(File.Exists);
        if (string.IsNullOrWhiteSpace(executablePath))
        {
            throw new InvalidOperationException("Chrome 또는 Edge 실행 파일을 찾지 못했습니다.");
        }

        var profileDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "KillRiceMonkey",
            Yes24ProfileName);
        Directory.CreateDirectory(profileDirectory);

        var startInfo = new ProcessStartInfo
        {
            FileName = executablePath,
            Arguments = $"--remote-debugging-port={Yes24RemoteDebugPort} --user-data-dir=\"{profileDirectory}\" --new-window \"{Yes24RemoteDebugLaunchUrl}\"",
            UseShellExecute = true
        };

        Process.Start(startInfo);

        var deadline = DateTimeOffset.UtcNow + TimeSpan.FromSeconds(10);
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (await PlaywrightRuntime.IsCdpEndpointAvailableAsync(Yes24CdpEndpoint, cancellationToken))
            {
                return $"remote debug 브라우저를 열었습니다: {Yes24RemoteDebugLaunchUrl}";
            }

            await Task.Delay(200, cancellationToken);
        }

        return $"브라우저 실행은 요청했지만 remote debug 포트 확인이 지연되고 있습니다. 직접 확인: {Yes24CdpEndpoint}";
    }

    public async Task<string> PrepareAutomationAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!await IsRemoteDebugBrowserAvailableAsync(cancellationToken))
        {
            throw new InvalidOperationException("먼저 YES24 Remote Debug 브라우저를 실행하세요.");
        }

        var page = await EnsurePreparedYes24ConnectedPageAsync(cancellationToken);
        if (page is null)
        {
            throw new InvalidOperationException("준비할 YES24 공연 페이지를 찾지 못했습니다. 공연 페이지를 연 뒤 다시 시도하세요.");
        }

        if (!await IsPreparedYes24PageReusableAsync(page, cancellationToken))
        {
            throw new InvalidOperationException("YES24 로그인 상태를 확인하지 못했습니다. YES24에 로그인한 공연 페이지에서 다시 시도하세요.");
        }

        await page.BringToFrontAsync();
        await EnsureYes24SalePopupClosedAsync(page);
        PlaywrightRuntime.EnsureDdddOcrWarmedUp();
        return $"YES24 준비 완료: {PlaywrightRuntime.SafePageUrl(page)}";
    }

    public async Task<bool> IsPageReadyAsync(CancellationToken cancellationToken)
    {
        if (await IsPreparedYes24PageReusableAsync(_preparedYes24Page, cancellationToken))
        {
            return true;
        }

        if (_preparedYes24ConnectedBrowser is null)
        {
            try
            {
                _preparedYes24ConnectedBrowser = await TryConnectToExistingYes24BrowserAsync(cancellationToken);
            }
            catch
            {
                return false;
            }

            if (_preparedYes24ConnectedBrowser is null)
            {
                return false;
            }
        }

        foreach (var page in _preparedYes24ConnectedBrowser.Contexts.SelectMany(x => x.Pages).Where(x => !x.IsClosed))
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (!PlaywrightRuntime.SafePageUrl(page).Contains("ticket.yes24.com/Perf/", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!await IsPreparedYes24PageReusableAsync(page, cancellationToken))
            {
                continue;
            }

            _preparedYes24Page = page;
            return true;
        }

        return false;
    }

    private async Task<AutomationRunResult> RunYes24AutomationAsync(TicketingJobRequest request, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        if (request.StepTimeoutSeconds <= 0)
        {
            return new AutomationRunResult(false, "단계별 제한 시간은 1초 이상이어야 합니다.", DateTimeOffset.Now);
        }

        if (!PlaywrightRuntime.TryParseDesiredDate(request.DesiredDate, out var desiredDate))
        {
            return new AutomationRunResult(false, "관람일 형식이 올바르지 않습니다. 예: 2026.04.11", DateTimeOffset.Now);
        }

        if (string.IsNullOrWhiteSpace(request.DesiredIdTime) && string.IsNullOrWhiteSpace(request.DesiredRound))
        {
            return new AutomationRunResult(false, "YES24 시간/회차 또는 IdTime 값이 필요합니다.", DateTimeOffset.Now);
        }

        var timeout = TimeSpan.FromSeconds(request.StepTimeoutSeconds);

        try
        {
            var cdpResult = await TryRunYes24AutomationViaConnectedBrowserAsync(request, desiredDate, timeout, progress, cancellationToken);
            if (cdpResult is not null)
            {
                return cdpResult;
            }

            return new AutomationRunResult(false, "YES24 remote debug 브라우저 연결 상태를 찾지 못했습니다.", DateTimeOffset.Now);
        }
        catch (TimeoutException ex)
        {
            _logger.LogError(ex, "[RunYes24] 시간 초과. date={Date}, round={Round}, idTime={IdTime}", desiredDate, request.DesiredRound, request.DesiredIdTime);
            return new AutomationRunResult(false, $"YES24 DOM 자동화 시간 초과: {ex.Message}", DateTimeOffset.Now);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[RunYes24] 예외 발생. date={Date}, round={Round}, idTime={IdTime}", desiredDate, request.DesiredRound, request.DesiredIdTime);
            return new AutomationRunResult(false, $"YES24 DOM 자동화 예외: {ex.Message}", DateTimeOffset.Now);
        }
    }

    private async Task<AutomationRunResult?> TryRunYes24AutomationViaConnectedBrowserAsync(
        TicketingJobRequest request,
        DateOnly desiredDate,
        TimeSpan timeout,
        IProgress<AutomationProgress>? progress,
        CancellationToken cancellationToken)
    {
        IPage? page = null;
        IPage? salePopup = null;

        try
        {
            page = await EnsurePreparedYes24ConnectedPageAsync(cancellationToken);
            if (page is null)
            {
                _logger.LogInformation("No existing YES24 browser with CDP endpoint was found.");
                return null;
            }

            await page.BringToFrontAsync();

            _logger.LogInformation("[YES24] 날짜 선택 시작. date={Date}", desiredDate);
            progress?.Report(new AutomationProgress("날짜 선택 중"));
            await SelectYes24DateAsync(page, desiredDate, timeout, cancellationToken);
            progress?.Report(new AutomationProgress("날짜 선택 완료", "날짜 선택 완료"));

            _logger.LogInformation("[YES24] 시간 선택 시작. round={Round}, idTime={IdTime}", request.DesiredRound, request.DesiredIdTime);
            progress?.Report(new AutomationProgress("시간 선택 중"));
            var idTime = await SelectYes24TimeAsync(page, request, timeout, cancellationToken);
            progress?.Report(new AutomationProgress("시간 선택 완료", $"시간 선택 완료: {idTime}"));

            _logger.LogInformation("[YES24] 예매 팝업 열기 시작. idTime={IdTime}", idTime);
            progress?.Report(new AutomationProgress("예매 클릭 중"));
            salePopup = await TriggerYes24BookingPopupAsync(page, idTime, timeout, progress, cancellationToken);
            _preparedYes24SalePopup = salePopup;
            progress?.Report(new AutomationProgress("예매 팝업 열림", "예매 팝업 열림"));

            salePopup.Dialog += async (_, dialog) =>
            {
                _logger.LogInformation("YES24 dialog 감지: type={Type}, message={Message}", dialog.Type, dialog.Message);
                try { await dialog.AcceptAsync(); } catch { }
                try { await salePopup.EvaluateAsync("() => { window.__yes24AlertDetected = true; }"); } catch { }
            };

            if (request.PauseBeforeSeatSelection && request.PauseGate is { } gate)
            {
                progress?.Report(new AutomationProgress("좌석 선택 대기 — 일시정지", "좌석 선택 전 일시정지됨. 재개 버튼을 눌러주세요."));
                await Task.Run(() => gate.Wait(cancellationToken), cancellationToken);
                progress?.Report(new AutomationProgress("좌석 선택 중", "일시정지 해제 — 좌석 선택 진행"));
            }

            _logger.LogInformation("[YES24] 좌석 iframe 로드 시작.");
            progress?.Report(new AutomationProgress("좌석 iframe 로드 중"));
            await TriggerYes24SeatLoadAsync(salePopup, timeout, cancellationToken);

            if (await TryHandleYes24StclabCaptchaAsync(salePopup, cancellationToken))
            {
                throw new InvalidOperationException("YES24 STCLAB CAPTCHA 감지 — 사용자 확인이 필요합니다.");
            }

            progress?.Report(new AutomationProgress("좌석 선택 중"));
            var seatTitle = await SelectYes24SeatAndAdvanceAsync(salePopup, request, timeout, progress, cancellationToken);
            progress?.Report(new AutomationProgress("좌석 선택 완료", $"YES24 좌석 선택 완료: {seatTitle}"));
            return new AutomationRunResult(true, $"YES24 좌석 선택 완료: {seatTitle}", DateTimeOffset.Now);
        }
        catch (Exception ex)
        {
            if (salePopup is not null && !salePopup.IsClosed)
            {
                _logger.LogWarning("[YES24] 실패 발생했으나 팝업은 닫지 않음 (재시도 시 재활용). popupUrl={Url}, exType={ExType}",
                    PlaywrightRuntime.SafePageUrl(salePopup), ex.GetType().Name);
            }

            if (page is not null)
            {
                _logger.LogError(ex, "[YES24] 자동화 실패. pageUrl={Url}", PlaywrightRuntime.SafePageUrl(page));
                throw new InvalidOperationException($"YES24 기존 브라우저 DOM 자동화 실패: {ex.Message}", ex);
            }

            _logger.LogWarning(ex, "Connected-browser YES24 automation failed before selecting a page.");
            return null;
        }
    }

    private async Task<IPage?> TryGetPreparedYes24PageAsync(CancellationToken cancellationToken)
    {
        await _yes24BrowserLock.WaitAsync(cancellationToken);
        try
        {
            if (await IsPreparedYes24PageReusableAsync(_preparedYes24Page, cancellationToken))
            {
                return _preparedYes24Page;
            }

            if (_preparedYes24ConnectedBrowser is not null)
            {
                var refreshedPage = await FindConnectedYes24PageAsync(_preparedYes24ConnectedBrowser, cancellationToken);
                if (await IsPreparedYes24PageReusableAsync(refreshedPage, cancellationToken))
                {
                    _preparedYes24Page = refreshedPage;
                    return refreshedPage;
                }
            }

            return null;
        }
        finally
        {
            _yes24BrowserLock.Release();
        }
    }

    private async Task<IPage?> EnsurePreparedYes24ConnectedPageAsync(CancellationToken cancellationToken)
    {
        await _yes24BrowserLock.WaitAsync(cancellationToken);
        try
        {
            if (await IsPreparedYes24PageReusableAsync(_preparedYes24Page, cancellationToken))
            {
                return _preparedYes24Page;
            }

            if (_preparedYes24ConnectedBrowser is not null)
            {
                var existingPage = await FindConnectedYes24PageAsync(_preparedYes24ConnectedBrowser, cancellationToken);
                if (await IsPreparedYes24PageReusableAsync(existingPage, cancellationToken))
                {
                    _preparedYes24Page = existingPage;
                    return existingPage;
                }
            }

            _preparedYes24ConnectedBrowser = await TryConnectToExistingYes24BrowserAsync(cancellationToken);
            if (_preparedYes24ConnectedBrowser is null)
            {
                _preparedYes24Page = null;
                return null;
            }

            _preparedYes24Page = await FindConnectedYes24PageAsync(_preparedYes24ConnectedBrowser, cancellationToken);
            return _preparedYes24Page;
        }
        finally
        {
            _yes24BrowserLock.Release();
        }
    }

    private async Task<bool> IsPreparedYes24PageReusableAsync(IPage? page, CancellationToken cancellationToken)
    {
        if (page is null || page.IsClosed)
        {
            return false;
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (!PlaywrightRuntime.SafePageUrl(page).Contains("ticket.yes24.com/Perf/", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return string.Equals(await TryGetYes24LoginStateAsync(page), "1", StringComparison.Ordinal);
    }

    private Task ReleasePreparedYes24ConnectionAsync()
    {
        _preparedYes24Page = null;
        _preparedYes24ConnectedBrowser = null;
        _preparedYes24SalePopup = null;
        return Task.CompletedTask;
    }

    private Task<IBrowser?> TryConnectToExistingYes24BrowserAsync(CancellationToken cancellationToken)
    {
        return _runtime.TryConnectToExistingChromiumBrowserAsync(Yes24CdpEndpoint, cancellationToken);
    }

    private async Task<IPage?> FindConnectedYes24PageAsync(IBrowser browser, CancellationToken cancellationToken)
    {
        var candidates = new List<(IPage Page, int Score)>();

        foreach (var page in browser.Contexts.SelectMany(x => x.Pages).Where(x => !x.IsClosed))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var url = PlaywrightRuntime.SafePageUrl(page);
            if (!url.Contains("ticket.yes24.com", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var score = 0;
            if (url.Contains("ticket.yes24.com", StringComparison.OrdinalIgnoreCase))
            {
                score += 2;
            }

            if (url.Contains("/Perf/", StringComparison.OrdinalIgnoreCase))
            {
                score += 6;
            }

            try
            {
                if (await page.Locator(".rn-04-left-calist a").CountAsync() > 0)
                {
                    score += 10;
                }

                if (string.Equals(await TryGetYes24LoginStateAsync(page), "1", StringComparison.Ordinal))
                {
                    score += 8;
                }

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

    private async Task<string> SelectYes24TimeAsync(IPage page, TicketingJobRequest request, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var timesLocator = page.Locator(".rn-04-left-calist a");

        await PlaywrightRuntime.WaitForConditionAsync(
            async () => await timesLocator.CountAsync() > 0,
            timeout,
            cancellationToken,
            "YES24 회차 목록을 찾지 못했습니다.");

        if (!string.IsNullOrWhiteSpace(request.DesiredIdTime))
        {
            var idTime = request.DesiredIdTime.Trim();
            var target = page.Locator($".rn-04-left-calist a[idTime='{idTime}']").First;
            if (await target.CountAsync() == 0)
            {
                throw new InvalidOperationException($"YES24 idTime을 찾지 못했습니다: {idTime}");
            }

            await target.ScrollIntoViewIfNeededAsync();
            await PlaywrightRuntime.ClickElementAsync(target);
            await PlaywrightRuntime.WaitForConditionAsync(
                async () => string.Equals(await target.GetAttributeAsync("idTime"), idTime, StringComparison.Ordinal)
                      && ((await target.GetAttributeAsync("class") ?? string.Empty).Contains("on", StringComparison.OrdinalIgnoreCase)
                          || string.Equals(await page.Locator(".rn-04-left-calist a.on").First.GetAttributeAsync("idTime"), idTime, StringComparison.Ordinal)),
                timeout,
                cancellationToken,
                "YES24 회차 선택 반영을 확인하지 못했습니다.");
            return idTime;
        }

        var selected = page.Locator(".rn-04-left-calist a.on").First;
        if (await selected.CountAsync() > 0)
        {
            var selectedIdTime = await selected.GetAttributeAsync("idTime");
            var selectedText = PlaywrightRuntime.NormalizeText(await selected.InnerTextAsync());
            if (string.IsNullOrWhiteSpace(request.DesiredRound) || IsMatchingYes24Time(selectedText, request.DesiredRound))
            {
                return selectedIdTime ?? throw new InvalidOperationException("YES24 선택된 회차의 idTime을 읽지 못했습니다.");
            }
        }

        var desiredRound = PlaywrightRuntime.NormalizeText(request.DesiredRound);
        var count = await timesLocator.CountAsync();
        for (var index = 0; index < count; index++)
        {
            var item = timesLocator.Nth(index);
            var timeText = PlaywrightRuntime.NormalizeText(await item.InnerTextAsync());
            if (!IsMatchingYes24Time(timeText, desiredRound))
            {
                continue;
            }

            var idTime = await item.GetAttributeAsync("idTime");
            if (string.IsNullOrWhiteSpace(idTime))
            {
                continue;
            }

            await item.ScrollIntoViewIfNeededAsync();
            await PlaywrightRuntime.ClickElementAsync(item);
            await PlaywrightRuntime.WaitForConditionAsync(
                async () => (await item.GetAttributeAsync("class") ?? string.Empty).Contains("on", StringComparison.OrdinalIgnoreCase)
                      || string.Equals(await page.Locator(".rn-04-left-calist a.on").First.GetAttributeAsync("idTime"), idTime, StringComparison.Ordinal),
                timeout,
                cancellationToken,
                "YES24 회차 선택 반영을 확인하지 못했습니다.");
            return idTime;
        }

        throw new InvalidOperationException($"YES24 시간을 찾지 못했습니다: {request.DesiredRound}");
    }

    private async Task<IPage> TriggerYes24BookingPopupAsync(IPage page, string idTime, TimeSpan timeout, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (_preparedYes24SalePopup is not null && !_preparedYes24SalePopup.IsClosed
            && PlaywrightRuntime.SafePageUrl(_preparedYes24SalePopup).Contains("/Pages/Perf/Sale/PerfSaleProcess.aspx", StringComparison.OrdinalIgnoreCase))
        {
            await _preparedYes24SalePopup.BringToFrontAsync();
            return _preparedYes24SalePopup;
        }

        var existingPopup = page.Context.Pages
            .FirstOrDefault(p => p != page && !p.IsClosed && PlaywrightRuntime.SafePageUrl(p).Contains("/Pages/Perf/Sale/PerfSaleProcess.aspx", StringComparison.OrdinalIgnoreCase));
        if (existingPopup is not null)
        {
            _preparedYes24SalePopup = existingPopup;
            await existingPopup.BringToFrontAsync();
            return existingPopup;
        }

        var beforePages = page.Context.Pages.Where(x => !x.IsClosed).ToHashSet();
        var idPerf = await ResolveYes24PerfIdAsync(page, cancellationToken);

        var directResult = await page.EvaluateAsync<string>("([perf, time]) => { if (typeof jsf_base_ShowPerfSaleProcess === 'function') { jsf_base_ShowPerfSaleProcess(perf, time); return 'ok'; } return 'missing'; }", new[] { idPerf, idTime });
        if (string.Equals(directResult, "ok", StringComparison.Ordinal))
        {
            var directPopup = await TryWaitForYes24SalePopupAsync(page, beforePages, TimeSpan.FromMilliseconds(350), cancellationToken);
            if (directPopup is not null)
            {
                _preparedYes24SalePopup = directPopup;
                await directPopup.BringToFrontAsync();
                return directPopup;
            }
        }

        var bookingButton = page.Locator("a.rn-bb03").First;
        if (await bookingButton.CountAsync() == 0)
        {
            throw new InvalidOperationException("YES24 예매 버튼을 찾지 못했습니다.");
        }

        await bookingButton.ScrollIntoViewIfNeededAsync();
        await PlaywrightRuntime.ClickBookingButtonAsync(bookingButton, timeout);

        var popup = await WaitForYes24SalePopupWithQueueAsync(page, timeout, progress, cancellationToken);
        if (popup is null)
        {
            throw new TimeoutException("YES24 예매 버튼 클릭 후 팝업 전환을 확인하지 못했습니다.");
        }

        _preparedYes24SalePopup = popup;
        await popup.BringToFrontAsync();
        return popup;
    }

    /// <summary>
    /// 예매하기 클릭 후 PerfSaleProcess.aspx 팝업 전환을 기다린다. 도중에 YES24 NetFunnel 대기열
    /// (`#NetFunnel_Skin_Top` 모달) 이 감지되면 취소 토큰이 내려올 때까지 무한 대기하면서 10초마다
    /// 진행 상황을 보고한다. 대기열이 감지되지 않은 상태로 timeout 이 만료되면 null 을 반환.
    /// </summary>
    private async Task<IPage?> WaitForYes24SalePopupWithQueueAsync(
        IPage page,
        TimeSpan timeout,
        IProgress<AutomationProgress>? progress,
        CancellationToken cancellationToken)
    {
        var queueEverDetected = false;
        var boundedDeadline = DateTimeOffset.UtcNow + timeout;

        // Phase 1: 기본 폴링 대기 (timeout 으로 제한). 도중 대기열 감지 시 즉시 Phase 2 로 진입.
        while (DateTimeOffset.UtcNow < boundedDeadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var target = FindYes24SalePopup(page);
            if (target is not null)
            {
                return target;
            }

            if (await IsYes24QueueActiveAsync(page))
            {
                queueEverDetected = true;
                break;
            }

            await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
        }

        // Phase 1.5: 타임아웃 직후에도 한 번 더 큐 상태를 재확인 (늦게 뜨는 경우 보호).
        if (!queueEverDetected)
        {
            queueEverDetected = await IsYes24QueueActiveAsync(page);
        }

        if (!queueEverDetected)
        {
            return null;
        }

        // Phase 2: 대기열 감지 — PerfSaleProcess.aspx 팝업 전환까지 무한 대기 (Melon/NOL 동일 패턴).
        _logger.LogInformation("[YES24] 대기열 감지 (#NetFunnel_Skin_Top) — PerfSaleProcess.aspx 전환까지 무한 대기");
        var queueSw = Stopwatch.StartNew();
        var lastQueueReportBucket = 0L;
        progress?.Report(new AutomationProgress("대기열 대기 중...", "YES24 NetFunnel 대기열 진입 — 팝업 전환 대기"));

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
                throw new InvalidOperationException("YES24 메인 페이지가 대기 중 닫혔습니다.");
            }

            var target = FindYes24SalePopup(page);
            if (target is not null)
            {
                _logger.LogInformation("[YES24] 대기열 종료 — PerfSaleProcess.aspx 도착. elapsed={Elapsed}s, url={Url}",
                    (int)queueSw.Elapsed.TotalSeconds, PlaywrightRuntime.SafePageUrl(target));
                progress?.Report(new AutomationProgress("대기열 통과", $"대기열 통과 ({(int)queueSw.Elapsed.TotalSeconds}초)"));
                return target;
            }

            await Task.Delay(100, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();
        return null;
    }

    private static IPage? FindYes24SalePopup(IPage page)
    {
        return page.Context.Pages
            .Where(x => !x.IsClosed)
            .FirstOrDefault(x => PlaywrightRuntime.SafePageUrl(x).Contains("/Pages/Perf/Sale/PerfSaleProcess.aspx", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// YES24 NetFunnel 대기열이 현재 활성 상태인지 확인.
    /// netfunnel-skin.js 의 yesticket 스킨은 `#NetFunnel_Skin_Top` 이라는 루트 div 를 inline 으로 삽입한다.
    /// 이 요소가 존재하고 실제 화면에 렌더링(visible)되어 있으면 대기열이 진행 중.
    /// </summary>
    private static async Task<bool> IsYes24QueueActiveAsync(IPage page)
    {
        try
        {
            return await page.EvaluateAsync<bool>(@"() => {
                const el = document.querySelector('#NetFunnel_Skin_Top');
                if (!el) return false;
                const style = window.getComputedStyle(el);
                if (style.display === 'none' || style.visibility === 'hidden') return false;
                const rect = el.getBoundingClientRect();
                return rect.width > 0 && rect.height > 0;
            }");
        }
        catch (PlaywrightException)
        {
            return false;
        }
    }

    private static async Task<IPage?> TryWaitForYes24SalePopupAsync(IPage page, HashSet<IPage> beforePages, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var openPages = page.Context.Pages.Where(x => !x.IsClosed).ToList();
            var newPopup = openPages.FirstOrDefault(x => !beforePages.Contains(x)
                && PlaywrightRuntime.SafePageUrl(x).Contains("/Pages/Perf/Sale/PerfSaleProcess.aspx", StringComparison.OrdinalIgnoreCase));
            if (newPopup is not null)
            {
                return newPopup;
            }

            var existingPopup = openPages.FirstOrDefault(x => PlaywrightRuntime.SafePageUrl(x).Contains("/Pages/Perf/Sale/PerfSaleProcess.aspx", StringComparison.OrdinalIgnoreCase));
            if (existingPopup is not null)
            {
                return existingPopup;
            }

            await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
        }

        return null;
    }

    private static async Task TriggerYes24SeatLoadAsync(IPage salePopup, TimeSpan timeout, CancellationToken cancellationToken)
    {
        await salePopup.BringToFrontAsync();

        // 1) 팝업 내부 alert 훅 설치 (2026-04-09 race condition 대응)
        //    fbk_Alert 는 `jgBookAlert === jcMODE_JQUERY (=1)` 일 때 native alert 가 아닌 `$j('#dialogAlert').jAlert(...)` HTML
        //    다이얼로그로 표시된다. Playwright 의 Page.Dialog 이벤트는 native 에만 반응하므로 JS 측 훅이 필요하다.
        //    window.alert 와 fbk_Alert 양쪽을 감싸서 `window.__yes24AlertDetected` / `window.__yes24LastAlert` 를 세팅.
        try
        {
            await salePopup.EvaluateAsync(@"() => {
                window.__yes24AlertDetected = false;
                window.__yes24LastAlert = null;
                if (!window.__yes24NativeAlertHooked) {
                    const origAlert = window.alert;
                    window.alert = function(message) {
                        window.__yes24AlertDetected = true;
                        window.__yes24LastAlert = 'alert:' + String(message);
                        if (typeof origAlert === 'function') {
                            return origAlert.apply(this, arguments);
                        }
                    };
                    window.__yes24NativeAlertHooked = true;
                }
                const hookFbk = () => {
                    if (typeof fbk_Alert === 'function' && !window.__yes24FbkHooked) {
                        const origFbk = fbk_Alert;
                        window.fbk_Alert = function(message) {
                            window.__yes24AlertDetected = true;
                            window.__yes24LastAlert = 'fbk:' + String(message);
                            return origFbk.apply(this, arguments);
                        };
                        window.__yes24FbkHooked = true;
                        return true;
                    }
                    return false;
                };
                if (!hookFbk()) {
                    const handle = setInterval(() => {
                        if (hookFbk()) {
                            clearInterval(handle);
                        }
                    }, 5);
                    setTimeout(() => { try { clearInterval(handle); } catch (e) {} }, 15000);
                }
            }");
        }
        catch (PlaywrightException)
        {
            // 팝업이 아직 실행 컨텍스트 준비 전이면 무시 — 다음 대기 루프에서 자연 해결
        }

        // 2) 팝업 초기화 대기 (2026-04-09 race condition 수정 핵심)
        //    `fdc_FlashSeatLoad` 내부 가드:
        //      if ($j("#IdTime").val() == "" || $j("#IdTime").val() == "0" || $j("#ulTime > li.on").length == 0) {
        //          fbk_Alert("공연회차를 선택하세요."); return;
        //      }
        //    `window.open` 직후 팝업이 "loading" → "interactive" → "complete" 전환 과정에서 약 520~617ms 구간에
        //    함수는 정의됐지만 #IdTime.value / #ulTime > li.on 이 아직 채워지지 않은 race window 가 존재한다.
        //    이 윈도우에서 호출하면 alert 가 발동하며, jQuery 다이얼로그라 Dialog 이벤트로도 잡히지 않는다.
        await PlaywrightRuntime.WaitForConditionAsync(
            async () =>
            {
                try
                {
                    return await salePopup.EvaluateAsync<bool>(@"() => {
                        if (typeof fdc_FlashSeatLoad !== 'function') return false;
                        const idInput = document.querySelector('#IdTime');
                        if (!idInput) return false;
                        const val = idInput.value || '';
                        if (val === '' || val === '0') return false;
                        return document.querySelectorAll('#ulTime > li.on').length > 0;
                    }");
                }
                catch (PlaywrightException)
                {
                    return false;
                }
            },
            timeout,
            cancellationToken,
            "YES24 예매 팝업 초기화(회차 정보 로드)를 확인하지 못했습니다.");

        // 3) alert 플래그 리셋 후 fdc_FlashSeatLoad 호출
        try { await salePopup.EvaluateAsync("() => { window.__yes24AlertDetected = false; window.__yes24LastAlert = null; }"); } catch { }

        var result = await salePopup.EvaluateAsync<string>(
            "() => { try { if (typeof fdc_FlashSeatLoad !== 'function') return 'missing'; fdc_FlashSeatLoad(); return 'ok'; } catch (e) { return 'err:' + (e && e.message ? e.message : e); } }");

        if (string.Equals(result, "missing", StringComparison.Ordinal))
        {
            var nextButton = salePopup.Locator("#StepCtrlBtn01 a").First;
            await PlaywrightRuntime.WaitForConditionAsync(
                async () => await nextButton.CountAsync() > 0,
                timeout,
                cancellationToken,
                "YES24 좌석 진입 버튼을 찾지 못했습니다.");
            await PlaywrightRuntime.ClickElementAsync(nextButton);
        }
        else if (!string.Equals(result, "ok", StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"YES24 fdc_FlashSeatLoad 실행 오류: {result}");
        }

        // 4) 호출 직후 alert 가 발동했는지 확인 — 초기화 가드를 통과했다면 정상적으로는 발생하지 않아야 한다.
        var alertAfter = await salePopup.EvaluateAsync<string>(
            "() => (window.__yes24AlertDetected === true ? String(window.__yes24LastAlert || '') : '')");
        if (!string.IsNullOrEmpty(alertAfter))
        {
            throw new InvalidOperationException($"YES24 fdc_FlashSeatLoad 호출 직후 알림 감지: {alertAfter}");
        }
    }

    private async Task<string> SelectYes24SeatAndAdvanceAsync(IPage salePopup, TicketingJobRequest request, TimeSpan timeout, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        const int maxSeatRetries = 10;
        var excludedSeats = new HashSet<string>(StringComparer.Ordinal);
        var desiredSeatIndex = Math.Max(request.DesiredSeatIndex, 1);

        for (var seatAttempt = 0; seatAttempt < maxSeatRetries; seatAttempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (salePopup.IsClosed)
            {
                throw new InvalidOperationException("YES24 예매 팝업이 닫혔습니다.");
            }

            try
            {
                var seatFrame = await FindYes24SeatFrameAsync(salePopup, timeout, cancellationToken);
                await EnsureYes24SeatInventoryReadyAsync(salePopup, seatFrame, request, timeout, cancellationToken);

                var clickResult = await ClickYes24SeatAsync(seatFrame, request.DesiredGrade, desiredSeatIndex, excludedSeats, cancellationToken);
                if (!string.Equals(clickResult.Status, "clicked", StringComparison.Ordinal))
                {
                    _logger.LogInformation("[YES24] 좌석 후보 없음 또는 클릭 실패. status={Status}, attempt={Attempt}/{Max}", clickResult.Status, seatAttempt + 1, maxSeatRetries);
                    await Task.Delay(50, cancellationToken);
                    continue;
                }

                var selectionVerified = await PlaywrightRuntime.TryWaitForConditionAsync(async () =>
                {
                    return await seatFrame.Locator("[name=tk].son").CountAsync() > 0
                        || await salePopup.Locator("#liSelSeat > li").CountAsync() > 0;
                }, TimeSpan.FromMilliseconds(300), cancellationToken);

                if (await DetectYes24SeatConflictAsync(salePopup))
                {
                    if (!string.IsNullOrWhiteSpace(clickResult.Value))
                    {
                        excludedSeats.Add(clickResult.Value);
                    }

                    await DismissYes24SeatConflictAlertAsync(salePopup);
                    progress?.Report(new AutomationProgress("좌석 재선택 중", "중복 좌석 감지 — 다른 좌석 재선택"));
                    continue;
                }

                if (!selectionVerified)
                {
                    if (!string.IsNullOrWhiteSpace(clickResult.Value))
                    {
                        excludedSeats.Add(clickResult.Value);
                    }

                    _logger.LogWarning("[YES24] 좌석 선택 검증 실패. seat={Seat}, attempt={Attempt}/{Max}", clickResult.Value, seatAttempt + 1, maxSeatRetries);
                    await Task.Delay(50, cancellationToken);
                    continue;
                }

                try
                {
                    await ConfirmYes24SeatAsync(salePopup, seatFrame, timeout, cancellationToken);
                    return string.IsNullOrWhiteSpace(clickResult.Title) ? (clickResult.Grade ?? clickResult.Value ?? "좌석") : clickResult.Title;
                }
                catch (Exception ex) when (ex is InvalidOperationException or TimeoutException)
                {
                    if (!string.IsNullOrWhiteSpace(clickResult.Value))
                    {
                        excludedSeats.Add(clickResult.Value);
                    }

                    await DismissYes24SeatConflictAlertAsync(salePopup);
                    _logger.LogWarning(ex, "[YES24] 좌석 선택 완료 진입 실패 — 다른 좌석으로 재시도. seat={Seat}, attempt={Attempt}/{Max}", clickResult.Value, seatAttempt + 1, maxSeatRetries);
                    progress?.Report(new AutomationProgress("좌석 재선택 중", "좌석 선택 완료 진입 실패 — 다른 좌석 재선택"));
                }
            }
            catch (PlaywrightException ex) when (seatAttempt < maxSeatRetries - 1)
            {
                _logger.LogWarning(ex, "[YES24] 좌석 선택 중 frame/page 재연결 필요. attempt={Attempt}/{Max}", seatAttempt + 1, maxSeatRetries);
                await Task.Delay(100, cancellationToken);
            }
        }

        throw new InvalidOperationException($"YES24 좌석 선택 실패 ({maxSeatRetries}회 시도).");
    }

    private async Task<IFrame> FindYes24SeatFrameAsync(IPage salePopup, TimeSpan timeout, CancellationToken cancellationToken)
    {
        foreach (var frame in salePopup.Frames)
        {
            if (frame == salePopup.MainFrame || !frame.Url.Contains("PerfSaleHtmlSeat.aspx", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            try
            {
                if (await frame.Locator("[name=tk]").CountAsync() > 0 || await frame.Locator("img[usemap]").CountAsync() > 0)
                {
                    return frame;
                }
            }
            catch (PlaywrightException)
            {
            }
        }

        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            foreach (var frame in salePopup.Frames)
            {
                if (frame == salePopup.MainFrame || !frame.Url.Contains("PerfSaleHtmlSeat.aspx", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    if (await frame.Locator("[name=tk]").CountAsync() > 0 || await frame.Locator("img[usemap]").CountAsync() > 0)
                    {
                        return frame;
                    }
                }
                catch (PlaywrightException)
                {
                }
            }

            await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
        }

        throw new TimeoutException("YES24 좌석 iframe을 찾지 못했습니다.");
    }

    private async Task EnsureYes24SeatInventoryReadyAsync(IPage salePopup, IFrame seatFrame, TicketingJobRequest request, TimeSpan timeout, CancellationToken cancellationToken)
    {
        if (await IsYes24ZoneSelectionRequiredAsync(seatFrame))
        {
            await EnsureYes24ZoneSelectedAsync(seatFrame, request, cancellationToken);
        }

        await PlaywrightRuntime.WaitForConditionAsync(async () =>
        {
            if (await TryHandleYes24StclabCaptchaAsync(salePopup, cancellationToken))
            {
                throw new InvalidOperationException("YES24 STCLAB CAPTCHA 감지 — 사용자 확인이 필요합니다.");
            }

            try
            {
                return await seatFrame.Locator("[name=tk]").CountAsync() > 0;
            }
            catch (PlaywrightException)
            {
                return false;
            }
        }, timeout, cancellationToken, "YES24 좌석 로드를 확인하지 못했습니다.");
    }

    private static async Task<bool> IsYes24ZoneSelectionRequiredAsync(IFrame seatFrame)
    {
        try
        {
            return await seatFrame.Locator("img[usemap]").CountAsync() > 0
                && await seatFrame.Locator("area[href^='javascript:ChangeBlock(']").CountAsync() > 0;
        }
        catch (PlaywrightException)
        {
            return false;
        }
    }

    private async Task EnsureYes24ZoneSelectedAsync(IFrame seatFrame, TicketingJobRequest request, CancellationToken cancellationToken)
    {
        if (!string.IsNullOrWhiteSpace(request.DesiredBlock))
        {
            await seatFrame.EvaluateAsync("(b) => ChangeBlock(parseInt(b, 10))", request.DesiredBlock);
            await PlaywrightRuntime.WaitForConditionAsync(
                async () => await seatFrame.Locator("[name=tk][grade]").CountAsync() > 0,
                TimeSpan.FromMilliseconds(500),
                cancellationToken,
                $"YES24 구역 좌석 로드를 확인하지 못했습니다: {request.DesiredBlock}");
            return;
        }

        var hrefs = await seatFrame.EvaluateAsync<string[]>("() => Array.from(document.querySelectorAll(\"area[href^='javascript:ChangeBlock('\")).map(area => area.getAttribute('href') || '')");
        foreach (var href in hrefs)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var match = ChangeBlockRegex.Match(href);
            if (!match.Success)
            {
                continue;
            }

            var blockId = match.Groups[1].Value;
            await seatFrame.EvaluateAsync("(b) => ChangeBlock(parseInt(b, 10))", blockId);
            var hasSeats = await PlaywrightRuntime.TryWaitForConditionAsync(
                async () => await seatFrame.Locator("[name=tk][grade]").CountAsync() > 0,
                TimeSpan.FromMilliseconds(500),
                cancellationToken);
            if (hasSeats)
            {
                return;
            }
        }

        throw new InvalidOperationException("YES24 구역 선택 후 가용 좌석을 찾지 못했습니다.");
    }

    private static async Task<Yes24SeatClickResult> ClickYes24SeatAsync(
        IFrame seatFrame,
        string? desiredGrade,
        int desiredSeatIndex,
        HashSet<string> excludedSeats,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var scanClickResult = await seatFrame.EvaluateAsync<string>(@"(args) => {
            const all = Array.from(document.querySelectorAll('[name=tk]'));
            const desiredGrade = args.desiredGrade || '';
            const desiredSeatIndex = Number(args.desiredSeatIndex || 1);
            const excluded = new Set(args.excluded || []);

            const available = all.filter(el => typeof el.onclick === 'function' || el.hasAttribute('grade'));
            const filtered = desiredGrade
                ? available.filter(el => ((el.getAttribute('grade') || '').includes(desiredGrade)))
                : available;

            filtered.sort((a, b) => {
                const at = a.title || '';
                const bt = b.title || '';
                if (at < bt) return -1;
                if (at > bt) return 1;
                const atop = parseInt(a.style.top || '0', 10) || 0;
                const btop = parseInt(b.style.top || '0', 10) || 0;
                if (atop !== btop) return atop - btop;
                const aleft = parseInt(a.style.left || '0', 10) || 0;
                const bleft = parseInt(b.style.left || '0', 10) || 0;
                return aleft - bleft;
            });

            if (filtered.length === 0) {
                return JSON.stringify({ s: 'empty' });
            }

            const candidates = filtered.filter(el => !excluded.has(el.getAttribute('value') || ''));
            if (candidates.length === 0) {
                return JSON.stringify({ s: 'not_found' });
            }

            const index = desiredSeatIndex > 0 && candidates.length >= desiredSeatIndex ? desiredSeatIndex - 1 : 0;
            const target = candidates[index] || candidates[0];
            if (!target) {
                return JSON.stringify({ s: 'not_found' });
            }

            target.click();
            return JSON.stringify({
                s: 'clicked',
                v: target.getAttribute('value') || '',
                t: target.title || '',
                g: target.getAttribute('grade') || ''
            });
        }", new
        {
            desiredGrade = desiredGrade?.Trim() ?? string.Empty,
            desiredSeatIndex,
            excluded = excludedSeats.ToArray()
        });

        using var doc = JsonDocument.Parse(scanClickResult);
        var root = doc.RootElement;
        return new Yes24SeatClickResult(
            root.GetProperty("s").GetString() ?? "not_found",
            root.TryGetProperty("v", out var value) ? value.GetString() : null,
            root.TryGetProperty("t", out var title) ? title.GetString() : null,
            root.TryGetProperty("g", out var grade) ? grade.GetString() : null);
    }

    private async Task ConfirmYes24SeatAsync(IPage salePopup, IFrame seatFrame, TimeSpan timeout, CancellationToken cancellationToken)
    {
        // YES24 좌석 선택 완료 흐름 (실측 검증):
        //   1) seat iframe 내부의 ChoiceEnd() 호출 → AJAX → ChoiceEnd_CallBack
        //   2) 응답 Code='none' 일 때 ChoiceEndProcess(token@class) → 부모 step01_time 의 selSeatClass 등 갱신
        //   3) 등급/매수 자동 선택 후 fdc_VerifySelSeatNumber() 가 step3 로 진입
        // - 응답 Code='block' 이면 STCLAB 캡차 → 상위에서 감지하여 사용자 개입 요청
        // - alert (다른 고객 결제중 등) 은 Dialog 핸들러에서 __yes24AlertDetected 로 표시 → 호출자가 재시도
        var iframeInvoke = await seatFrame.EvaluateAsync<string>(
            "() => { try { if (typeof ChoiceEnd === 'function') { ChoiceEnd(); return 'ok'; } return 'missing'; } catch (e) { return 'err:' + e.message; } }");
        if (string.Equals(iframeInvoke, "missing", StringComparison.Ordinal))
        {
            // 폴백: iframe 내부 '좌석선택완료' 링크를 직접 클릭
            var endLink = seatFrame.Locator("a[href*='ChoiceEnd']").First;
            if (await endLink.CountAsync() == 0)
            {
                throw new InvalidOperationException("YES24 좌석 선택 완료(ChoiceEnd) 진입 수단을 찾지 못했습니다.");
            }

            await PlaywrightRuntime.ClickElementAsync(endLink);
        }
        else if (!string.Equals(iframeInvoke, "ok", StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"YES24 ChoiceEnd 호출 실패: {iframeInvoke}");
        }

        // ChoiceEnd_CallBack 후 부모 step01_time 의 selSeatClass 가 등장하면 등급/수량 자동 확정 후 fdc_VerifySelSeatNumber 로 step3 전환
        var preconditionMet = await PlaywrightRuntime.TryWaitForConditionAsync(async () =>
        {
            if (salePopup.IsClosed)
            {
                return false;
            }

            if (await DetectYes24SeatConflictAsync(salePopup))
            {
                return false;
            }

            return await salePopup.EvaluateAsync<bool>(@"() => {
                const sel = document.querySelector(""#step01_time #ulSeatSpace select[id='selSeatClass']"");
                return !!sel;
            }");
        }, timeout, cancellationToken);

        if (await DetectYes24SeatConflictAsync(salePopup))
        {
            throw new InvalidOperationException("YES24 좌석 중복 감지");
        }

        if (await TryHandleYes24StclabCaptchaAsync(salePopup, cancellationToken))
        {
            throw new InvalidOperationException("YES24 STCLAB CAPTCHA 감지 — 사용자 확인이 필요합니다.");
        }

        if (preconditionMet)
        {
            await salePopup.EvaluateAsync(@"() => {
                const selects = document.querySelectorAll(""#step01_time #ulSeatSpace select[id='selSeatClass']"");
                selects.forEach(sel => {
                    let picked = null;
                    for (const opt of sel.options) {
                        if (opt.value && opt.value !== '0' && opt.value !== '-1') { picked = opt.value; break; }
                    }
                    if (picked) {
                        sel.value = picked;
                        sel.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                });
            }");
        }

        var verifyResult = await salePopup.EvaluateAsync<string>(
            "() => { try { if (typeof fdc_VerifySelSeatNumber === 'function') { fdc_VerifySelSeatNumber(); return 'ok'; } return 'missing'; } catch (e) { return 'err:' + e.message; } }");
        if (string.Equals(verifyResult, "missing", StringComparison.Ordinal))
        {
            var nextButton = salePopup.Locator("#StepCtrlBtn01 a").First;
            if (await nextButton.CountAsync() > 0)
            {
                await PlaywrightRuntime.ClickElementAsync(nextButton);
            }
        }
        else if (!string.Equals(verifyResult, "ok", StringComparison.Ordinal))
        {
            _logger.LogWarning("[YES24] fdc_VerifySelSeatNumber 호출 결과: {Result}", verifyResult);
        }

        var advanced = await PlaywrightRuntime.TryWaitForConditionAsync(async () =>
        {
            if (salePopup.IsClosed)
            {
                return false;
            }

            return await salePopup.EvaluateAsync<bool>(@"() => {
                const step01 = document.querySelector('#step01');
                const step03 = document.querySelector('#step03');
                if (!step01 || !step03) {
                    return false;
                }

                const step01Display = window.getComputedStyle(step01).display;
                const step03Display = window.getComputedStyle(step03).display;
                return step01Display === 'none' && step03Display === 'block';
            }");
        }, timeout, cancellationToken);

        if (advanced)
        {
            return;
        }

        if (await DetectYes24SeatConflictAsync(salePopup))
        {
            throw new InvalidOperationException("YES24 좌석 중복 감지");
        }

        if (await TryHandleYes24StclabCaptchaAsync(salePopup, cancellationToken))
        {
            throw new InvalidOperationException("YES24 STCLAB CAPTCHA 감지 — 사용자 확인이 필요합니다.");
        }

        throw new TimeoutException("YES24 step01 → step03 전환을 확인하지 못했습니다.");
    }

    private static async Task<bool> DetectYes24SeatConflictAsync(IPage salePopup)
    {
        try
        {
            return await salePopup.EvaluateAsync<bool>("() => window.__yes24AlertDetected === true");
        }
        catch (PlaywrightException)
        {
            return false;
        }
    }

    private static async Task DismissYes24SeatConflictAlertAsync(IPage salePopup)
    {
        try
        {
            await salePopup.EvaluateAsync("() => { window.__yes24AlertDetected = false; }");
        }
        catch (PlaywrightException)
        {
        }
    }

    private async Task<bool> TryHandleYes24StclabCaptchaAsync(IPage salePopup, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        try
        {
            if (PlaywrightRuntime.SafePageUrl(salePopup).Contains("cdn-botmanager.stclab.com/yes24/", StringComparison.OrdinalIgnoreCase)
                || salePopup.Frames.Any(frame => frame.Url.Contains("cdn-botmanager.stclab.com/yes24/", StringComparison.OrdinalIgnoreCase)))
            {
                _logger.LogWarning("YES24 STCLAB CAPTCHA 감지. popupUrl={Url}", PlaywrightRuntime.SafePageUrl(salePopup));
                return true;
            }
        }
        catch (PlaywrightException)
        {
        }

        // TODO: future STCLAB CAPTCHA handling hook.
        return false;
    }

    private static async Task EnsureYes24SalePopupClosedAsync(IPage mainPage)
    {
        var popups = mainPage.Context.Pages
            .Where(page => page != mainPage
                && !page.IsClosed
                && PlaywrightRuntime.SafePageUrl(page).Contains("/Pages/Perf/Sale/PerfSaleProcess.aspx", StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var popup in popups)
        {
            try
            {
                await popup.CloseAsync();
            }
            catch (PlaywrightException)
            {
            }
        }
    }

    private static async Task SelectYes24DateAsync(IPage page, DateOnly desiredDate, TimeSpan timeout, CancellationToken cancellationToken)
    {
        await PlaywrightRuntime.WaitForConditionAsync(
            async () => await page.Locator("#rncalendar").CountAsync() > 0,
            timeout,
            cancellationToken,
            "YES24 달력을 찾지 못했습니다.");

        if (await IsYes24DateSelectedAsync(page, desiredDate))
        {
            return;
        }

        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var clickResult = await page.EvaluateAsync<string>(@"(args) => {
                const cells = Array.from(document.querySelectorAll('#rncalendar td[data-year][data-month]'));
                const target = cells.find(td =>
                    td.getAttribute('data-year') === String(args.year)
                    && td.getAttribute('data-month') === String(args.month)
                    && (td.querySelector('a')?.textContent || '').trim() === String(args.day));

                if (!target) {
                    return 'missing';
                }

                const anchor = target.querySelector('a');
                if (!anchor) {
                    return 'missing';
                }

                anchor.click();
                return 'clicked';
            }", new
            {
                year = desiredDate.Year,
                month = desiredDate.Month - 1,
                day = desiredDate.Day
            });

            if (string.Equals(clickResult, "clicked", StringComparison.Ordinal))
            {
                await PlaywrightRuntime.WaitForConditionAsync(
                    async () => await IsYes24DateSelectedAsync(page, desiredDate),
                    timeout,
                    cancellationToken,
                    "YES24 관람일 선택 반영을 확인하지 못했습니다.");
                return;
            }

            var displayedMonth = await page.EvaluateAsync<string>(@"() => {
                const first = document.querySelector('#rncalendar td[data-year][data-month]');
                if (!first) {
                    return '';
                }

                return `${first.getAttribute('data-year')}-${Number(first.getAttribute('data-month')) + 1}`;
            }");

            if (string.IsNullOrWhiteSpace(displayedMonth))
            {
                await Task.Delay(50, cancellationToken);
                continue;
            }

            var parts = displayedMonth.Split('-', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 2 || !int.TryParse(parts[0], out var displayedYear) || !int.TryParse(parts[1], out var displayedMonthValue))
            {
                await Task.Delay(50, cancellationToken);
                continue;
            }

            var currentMonth = new DateOnly(displayedYear, displayedMonthValue, 1);
            var desiredMonth = new DateOnly(desiredDate.Year, desiredDate.Month, 1);
            var moveSelector = desiredMonth >= currentMonth ? "#rncalendar .ui-datepicker-next" : "#rncalendar .ui-datepicker-prev";
            var moveButton = page.Locator(moveSelector).First;
            if (await moveButton.CountAsync() == 0)
            {
                break;
            }

            await PlaywrightRuntime.ClickElementAsync(moveButton);
            await Task.Delay(50, cancellationToken);
        }

        throw new InvalidOperationException($"YES24 관람일을 찾지 못했습니다: {desiredDate:yyyy.MM.dd}");
    }

    private static async Task<bool> IsYes24DateSelectedAsync(IPage page, DateOnly desiredDate)
    {
        try
        {
            return await page.EvaluateAsync<bool>(@"(args) => {
                const current = document.querySelector('#rncalendar td.ui-datepicker-current-day');
                if (!current) {
                    return false;
                }

                const anchor = current.querySelector('a');
                return current.getAttribute('data-year') === String(args.year)
                    && current.getAttribute('data-month') === String(args.month)
                    && (anchor?.textContent || '').trim() === String(args.day);
            }", new
            {
                year = desiredDate.Year,
                month = desiredDate.Month - 1,
                day = desiredDate.Day
            });
        }
        catch (PlaywrightException)
        {
            return false;
        }
    }

    private static async Task<string> ResolveYes24PerfIdAsync(IPage page, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        try
        {
            var fromPage = await page.EvaluateAsync<string>("() => { const hidden = document.querySelector('#HidIdPerf'); return hidden ? String(hidden.value || '').trim() : ''; }");
            if (!string.IsNullOrWhiteSpace(fromPage))
            {
                return fromPage;
            }
        }
        catch (PlaywrightException)
        {
        }

        var match = PerfIdRegex.Match(PlaywrightRuntime.SafePageUrl(page));
        if (match.Success)
        {
            return match.Groups["id"].Value;
        }

        throw new InvalidOperationException("YES24 공연 ID(IdPerf)를 확인하지 못했습니다.");
    }

    private static async Task<string> TryGetYes24LoginStateAsync(IPage page)
    {
        try
        {
            return await page.EvaluateAsync<string>("() => (typeof IsLogin !== 'undefined' ? String(IsLogin) : '')");
        }
        catch (PlaywrightException)
        {
            return string.Empty;
        }
    }

    private static bool IsMatchingYes24Time(string actual, string desired)
    {
        var normalizedActual = PlaywrightRuntime.NormalizeText(actual);
        var normalizedDesired = PlaywrightRuntime.NormalizeText(desired);
        if (string.IsNullOrWhiteSpace(normalizedDesired))
        {
            return false;
        }

        if (string.Equals(normalizedActual, normalizedDesired, StringComparison.OrdinalIgnoreCase)
            || normalizedActual.Contains(normalizedDesired, StringComparison.OrdinalIgnoreCase)
            || normalizedDesired.Contains(normalizedActual, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (TryExtractYes24TimeValue(normalizedActual, out var actualTime)
            && TryExtractYes24TimeValue(normalizedDesired, out var desiredTime))
        {
            return actualTime == desiredTime;
        }

        var actualDigits = PlaywrightRuntime.DigitsOnlyPattern.Replace(normalizedActual, string.Empty);
        var desiredDigits = PlaywrightRuntime.DigitsOnlyPattern.Replace(normalizedDesired, string.Empty);
        if (!string.IsNullOrWhiteSpace(actualDigits) && !string.IsNullOrWhiteSpace(desiredDigits))
        {
            return actualDigits.EndsWith(desiredDigits, StringComparison.Ordinal)
                || desiredDigits.EndsWith(actualDigits, StringComparison.Ordinal);
        }

        return false;
    }

    private static bool TryExtractYes24TimeValue(string text, out int timeValue)
    {
        timeValue = default;

        var koreanMatch = Regex.Match(text, @"(?<period>오전|오후)?\s*(?<hour>\d{1,2})\s*(?:시|:)\s*(?<minute>\d{1,2})");
        if (!koreanMatch.Success)
        {
            return false;
        }

        if (!int.TryParse(koreanMatch.Groups["hour"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var hour)
            || !int.TryParse(koreanMatch.Groups["minute"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var minute))
        {
            return false;
        }

        var period = koreanMatch.Groups["period"].Value;
        if (string.Equals(period, "오후", StringComparison.Ordinal) && hour < 12)
        {
            hour += 12;
        }
        else if (string.Equals(period, "오전", StringComparison.Ordinal) && hour == 12)
        {
            hour = 0;
        }

        timeValue = (hour * 100) + minute;
        return true;
    }

    private sealed record Yes24SeatClickResult(string Status, string? Value, string? Title, string? Grade);
}
