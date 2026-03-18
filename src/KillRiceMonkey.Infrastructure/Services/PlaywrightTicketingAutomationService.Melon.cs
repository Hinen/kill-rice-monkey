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

public sealed partial class PlaywrightTicketingAutomationService
{
    private async Task<AutomationRunResult> RunMelonAutomationAsync(TicketingJobRequest request, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        if (request.StepTimeoutSeconds <= 0)
        {
            return new AutomationRunResult(false, "단계별 제한 시간은 1초 이상이어야 합니다.", DateTimeOffset.Now);
        }

        if (!TryParseDesiredDate(request.DesiredDate, out var desiredDate))
        {
            return new AutomationRunResult(false, "관람일 형식이 올바르지 않습니다. 예: 2026.03.14", DateTimeOffset.Now);
        }

        if (string.IsNullOrWhiteSpace(request.DesiredRound))
        {
            return new AutomationRunResult(false, "시간 값이 비어 있습니다. 예: 18:00", DateTimeOffset.Now);
        }

        var desiredTime = NormalizeText(request.DesiredRound);
        var timeout = TimeSpan.FromSeconds(request.StepTimeoutSeconds);

        try
        {
            var cdpResult = await TryRunMelonAutomationViaConnectedBrowserAsync(request, desiredDate, desiredTime, timeout, progress, cancellationToken);
            if (cdpResult is not null)
            {
                return cdpResult;
            }

            return new AutomationRunResult(false, "Melon remote debug 브라우저 연결 상태를 찾지 못했습니다.", DateTimeOffset.Now);
        }
        catch (TimeoutException ex)
        {
            _logger.LogError(ex, "[RunMelon] 시간 초과. date={Date}, time={Time}, exType={ExType}", desiredDate, desiredTime, ex.GetType().Name);
            return new AutomationRunResult(false, $"Melon DOM 자동화 시간 초과: {ex.Message}", DateTimeOffset.Now);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[RunMelon] 예외 발생. date={Date}, time={Time}, exType={ExType}, innerEx={InnerExType}",
                desiredDate, desiredTime, ex.GetType().Name, ex.InnerException?.GetType().Name ?? "없음");
            return new AutomationRunResult(false, $"Melon DOM 자동화 예외: {ex.Message}", DateTimeOffset.Now);
        }
    }

    private async Task<AutomationRunResult?> TryRunMelonAutomationViaConnectedBrowserAsync(TicketingJobRequest request, DateOnly desiredDate, string desiredTime, TimeSpan timeout, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        IPage? page = null;
        IPage? captchaPage = null;
        try
        {
            page = await EnsurePreparedMelonConnectedPageAsync(cancellationToken);
            if (page is null)
            {
                _logger.LogInformation("No existing Melon browser with CDP endpoint was found.");
                return null;
            }

            _logger.LogInformation("Connected-browser Melon automation selected page. url={Url}", SafePageUrl(page));
            await page.BringToFrontAsync();
            if (!_melonPopupClosedDuringPrepare)
            {
                await EnsureMelonPopupClosedAsync(page, TimeSpan.FromMilliseconds(500), cancellationToken);
            }

            _melonPopupClosedDuringPrepare = false;

            _logger.LogInformation("[Melon] 날짜 선택 시작. date={Date}", desiredDate);
            progress?.Report(new AutomationProgress("날짜 선택 중"));
            await SelectMelonDateAsync(page, desiredDate, timeout, cancellationToken);
            _logger.LogInformation("[Melon] 날짜 선택 완료.");
            progress?.Report(new AutomationProgress("날짜 선택 완료", "날짜 선택 완료"));

            _logger.LogInformation("[Melon] 시간 선택 시작. time={Time}", desiredTime);
            progress?.Report(new AutomationProgress("시간 선택 중"));
            await SelectMelonTimeAsync(page, desiredTime, timeout, cancellationToken);
            _logger.LogInformation("[Melon] 시간 선택 완료.");
            progress?.Report(new AutomationProgress("시간 선택 완료", "시간 선택 완료"));

            _logger.LogInformation("[Melon] 예매하기 버튼 클릭 시작.");
            progress?.Report(new AutomationProgress("예매 클릭 중"));
            captchaPage = await ClickMelonBookingAsync(page, timeout, progress, cancellationToken);
            _logger.LogInformation("[Melon] 예매 팝업 열림. popupUrl={Url}", SafePageUrl(captchaPage));
            progress?.Report(new AutomationProgress("예매 팝업 열림", "예매 팝업 열림 (대기열 포함 가능)"));

            captchaPage.Dialog += async (_, dialog) =>
            {
                _logger.LogInformation("멜론 dialog 감지: type={Type}, message={Message}", dialog.Type, dialog.Message);
                try { await dialog.AcceptAsync(); } catch { }
                try { await captchaPage.EvaluateAsync("() => { window.__melonAlertDetected = true; }"); } catch { }
            };

            var captchaSw = Stopwatch.StartNew();
            _logger.LogInformation("[Melon] CAPTCHA 풀이 시작.");
            progress?.Report(new AutomationProgress("캡차 입력 중"));
            await SolveMelonCaptchaAsync(captchaPage, timeout, cancellationToken);
            _logger.LogInformation("[PERF] SolveCaptcha: {Ms}ms. popupUrl={Url}, isClosed={IsClosed}",
                captchaSw.ElapsedMilliseconds, SafePageUrl(captchaPage), captchaPage.IsClosed);
            progress?.Report(new AutomationProgress("캡차 입력 완료", "캡차 처리 완료"));

            if (captchaPage.IsClosed)
            {
                _logger.LogWarning("[Melon] CAPTCHA 풀이 후 팝업이 닫혀 있음 — 재시도 필요.");
                return new AutomationRunResult(false, "CAPTCHA 풀이 후 팝업 닫힘 — 재시도.", DateTimeOffset.Now);
            }

            const int maxSeatRetries = 10;
            for (var seatAttempt = 0; seatAttempt < maxSeatRetries; seatAttempt++)
            {
                try
                {
                    _logger.LogInformation("[Melon] 좌석 선택 시도 {Attempt}/{Max}. popupUrl={Url}, frameCount={FrameCount}",
                        seatAttempt + 1, maxSeatRetries, SafePageUrl(captchaPage), captchaPage.Frames.Count);

                    if (seatAttempt == 0 && request.PauseGate is { } gate)
                    {
                        _logger.LogInformation("[Melon] 좌석 선택 전 일시정지 — 사용자 재개 대기 중.");
                        progress?.Report(new AutomationProgress("좌석 선택 대기 — 일시정지", "좌석 선택 전 일시정지됨. 재개 버튼을 눌러주세요."));
                        await Task.Run(() => gate.Wait(cancellationToken), cancellationToken);
                        _logger.LogInformation("[Melon] 일시정지 해제 — 좌석 선택 진행.");
                        progress?.Report(new AutomationProgress("좌석 선택 중", "일시정지 해제 — 좌석 선택 진행"));
                    }

                    progress?.Report(new AutomationProgress("좌석 선택 중"));
                    await SelectMelonSeatAndCompleteAsync(captchaPage, timeout, progress, cancellationToken);
                    _logger.LogInformation("[Melon] 좌석 선택 및 완료 버튼 클릭 성공!");
                    progress?.Report(new AutomationProgress("좌석 선택 완료", "좌석 선택 및 완료 버튼 클릭"));
                    return new AutomationRunResult(true, $"Melon 기존 브라우저 DOM 자동화 완료: {desiredDate:yyyy.MM.dd} / {desiredTime} 선택, 좌석 선택 완료.", DateTimeOffset.Now);
                }
                catch (Exception seatEx) when (seatAttempt < maxSeatRetries - 1)
                {
                    _logger.LogWarning(seatEx, "[Melon] 좌석 선택 실패 (attempt={Attempt}/{Max}). 같은 팝업에서 재시도합니다. popupUrl={Url}, isClosed={IsClosed}",
                        seatAttempt + 1, maxSeatRetries, SafePageUrl(captchaPage), captchaPage.IsClosed);

                    if (captchaPage.IsClosed)
                    {
                        _logger.LogWarning("[Melon] 팝업이 닫혀 있어 좌석 재시도 불가.");
                        break;
                    }

                    await Task.Delay(100, cancellationToken);
                }
            }

            _logger.LogError("[Melon] 좌석 선택 {Max}회 모두 실패. 전체 재시도로 전환.", maxSeatRetries);
            throw new InvalidOperationException($"좌석 선택 {maxSeatRetries}회 시도 모두 실패.");
        }
        catch (Exception ex)
        {
            if (captchaPage is not null && !captchaPage.IsClosed)
            {
                _logger.LogWarning("[Melon] 실패 발생했으나 팝업은 닫지 않음 (재시도 시 재활용). popupUrl={Url}, exType={ExType}",
                    SafePageUrl(captchaPage), ex.GetType().Name);
            }

            if (page is not null)
            {
                _logger.LogError(ex, "[Melon] 자동화 실패. 예외를 상위로 전파합니다. pageUrl={Url}", SafePageUrl(page));
                throw new InvalidOperationException($"Melon 기존 브라우저 DOM 자동화 실패: {ex.Message}", ex);
            }

            _logger.LogWarning(ex, "Connected-browser Melon automation failed before selecting a page.");
            return null;
        }
    }

    private async Task CloseMelonPopupPagesSafelyAsync(IPage? mainPage, IPage? captchaPage)
    {
        try
        {
            if (captchaPage is not null && !captchaPage.IsClosed)
            {
                _logger.LogInformation("[Melon] onestop 팝업 닫기: url={Url}", SafePageUrl(captchaPage));
                await captchaPage.CloseAsync();
            }
        }
        catch (PlaywrightException ex)
        {
            _logger.LogWarning(ex, "[Melon] 팝업 닫기 실패 (무시).");
        }

        if (mainPage is null) return;
        try
        {
            var remainingPopups = mainPage.Context.Pages
                .Where(p => p != mainPage && !p.IsClosed && SafePageUrl(p).Contains("/reservation/popup/onestop.htm", StringComparison.OrdinalIgnoreCase))
                .ToList();
            foreach (var popup in remainingPopups)
            {
                _logger.LogInformation("[Melon] 잔여 팝업 닫기: url={Url}", SafePageUrl(popup));
                try { await popup.CloseAsync(); } catch (PlaywrightException) { }
            }
        }
        catch (PlaywrightException) { }
    }

    private async Task<IPage?> TryGetPreparedMelonPageAsync(CancellationToken cancellationToken)
    {
        await _melonBrowserLock.WaitAsync(cancellationToken);
        try
        {
            if (await IsPreparedMelonPageReusableAsync(_preparedMelonPage, cancellationToken))
            {
                return _preparedMelonPage;
            }

            if (_preparedMelonConnectedBrowser is not null)
            {
                var refreshedPage = await FindConnectedMelonPageAsync(_preparedMelonConnectedBrowser, cancellationToken);
                if (await IsPreparedMelonPageReusableAsync(refreshedPage, cancellationToken))
                {
                    _preparedMelonPage = refreshedPage;
                    return refreshedPage;
                }
            }

            return null;
        }
        finally
        {
            _melonBrowserLock.Release();
        }
    }

    private async Task<IPage?> EnsurePreparedMelonConnectedPageAsync(CancellationToken cancellationToken)
    {
        await _melonBrowserLock.WaitAsync(cancellationToken);
        try
        {
            if (await IsPreparedMelonPageReusableAsync(_preparedMelonPage, cancellationToken))
            {
                return _preparedMelonPage;
            }

            if (_preparedMelonConnectedBrowser is not null)
            {
                var existingPage = await FindConnectedMelonPageAsync(_preparedMelonConnectedBrowser, cancellationToken);
                if (await IsPreparedMelonPageReusableAsync(existingPage, cancellationToken))
                {
                    _preparedMelonPage = existingPage;
                    return existingPage;
                }
            }

            _playwright ??= await Playwright.CreateAsync();
            _preparedMelonConnectedBrowser = await TryConnectToExistingMelonBrowserAsync(_playwright, cancellationToken);
            if (_preparedMelonConnectedBrowser is null)
            {
                _preparedMelonPage = null;
                return null;
            }

            _preparedMelonPage = await FindConnectedMelonPageAsync(_preparedMelonConnectedBrowser, cancellationToken);
            return _preparedMelonPage;
        }
        finally
        {
            _melonBrowserLock.Release();
        }
    }

    private static Task<bool> IsPreparedMelonPageReusableAsync(IPage? page, CancellationToken cancellationToken)
    {
        if (page is null || page.IsClosed)
        {
            return Task.FromResult(false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(SafePageUrl(page).Contains("ticket.melon.com/performance/index.htm", StringComparison.OrdinalIgnoreCase));
    }

    private Task ReleasePreparedMelonConnectionAsync()
    {
        _preparedMelonPage = null;
        _preparedMelonConnectedBrowser = null;
        _melonPopupClosedDuringPrepare = false;
        return Task.CompletedTask;
    }

    public async Task<bool> IsMelonPageReadyAsync(CancellationToken cancellationToken)
    {
        if (_preparedMelonPage is not null && !_preparedMelonPage.IsClosed)
        {
            if (SafePageUrl(_preparedMelonPage).Contains("ticket.melon.com/performance/index.htm", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        if (_preparedMelonConnectedBrowser is null)
        {
            try
            {
                _playwright ??= await Playwright.CreateAsync();
                _preparedMelonConnectedBrowser = await TryConnectToExistingMelonBrowserAsync(_playwright, cancellationToken);
            }
            catch
            {
                return false;
            }

            if (_preparedMelonConnectedBrowser is null)
            {
                return false;
            }
        }

        foreach (var page in _preparedMelonConnectedBrowser.Contexts.SelectMany(x => x.Pages).Where(x => !x.IsClosed))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var url = SafePageUrl(page);
            if (!url.Contains("ticket.melon.com/performance/index.htm", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            _preparedMelonPage = page;
            return true;
        }

        return false;
    }

    private static Task<IBrowser?> TryConnectToExistingMelonBrowserAsync(IPlaywright playwright, CancellationToken cancellationToken)
    {
        return TryConnectToExistingChromiumBrowserAsync(playwright, MelonCdpEndpoint, cancellationToken);
    }
    private static async Task<IPage?> FindConnectedMelonPageAsync(IBrowser browser, CancellationToken cancellationToken)
    {
        var candidates = new List<(IPage Page, int Score)>();
        foreach (var page in browser.Contexts.SelectMany(x => x.Pages).Where(x => !x.IsClosed))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var url = SafePageUrl(page);
            var score = 0;
            if (url.Contains("ticket.melon.com", StringComparison.OrdinalIgnoreCase))
            {
                score += 2;
            }

            if (url.Contains("/performance/index.htm", StringComparison.OrdinalIgnoreCase))
            {
                score += 4;
            }

            if (await TryWaitForConditionAsync(async () => await page.Locator("#ticketing_process_box .wrap_ticketing_process").CountAsync() > 0, TimeSpan.FromMilliseconds(150), cancellationToken))
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
    private async Task SelectMelonSeatAndCompleteAsync(IPage page, TimeSpan timeout, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        var totalSw = Stopwatch.StartNew();
        _logger.LogInformation("[SelectSeat] 멜론 좌석 선택 시작. url={Url}, isClosed={IsClosed}, frameCount={FrameCount}",
            SafePageUrl(page), page.IsClosed, page.Frames.Count);

        var stepSw = Stopwatch.StartNew();
        var seatFrame = await FindMelonSeatFrameAsync(page, timeout, cancellationToken);
        _logger.LogInformation("[PERF] FindMelonSeatFrame: {Ms}ms. frameUrl={Url}", stepSw.ElapsedMilliseconds, seatFrame.Url);

        stepSw.Restart();
        var zoneRequired = await IsMelonZoneSelectionRequiredAsync(seatFrame);
        _logger.LogInformation("[PERF] IsMelonZoneSelectionRequired: {Ms}ms. required={Required}", stepSw.ElapsedMilliseconds, zoneRequired);
        if (zoneRequired)
        {
            stepSw.Restart();
            await WaitForMelonZoneSelectionAsync(seatFrame, progress, cancellationToken);
            _logger.LogInformation("[PERF] WaitForMelonZoneSelection: {Ms}ms", stepSw.ElapsedMilliseconds);
        }

        stepSw.Restart();
        var validFrame = await SelectMelonSeatInFrameAsync(page, seatFrame, timeout, cancellationToken);
        _logger.LogInformation("[PERF] SelectMelonSeatInFrame: {Ms}ms", stepSw.ElapsedMilliseconds);

        stepSw.Restart();
        await ClickMelonSeatCompleteAsync(page, validFrame, timeout, cancellationToken);
        _logger.LogInformation("[PERF] ClickMelonSeatComplete: {Ms}ms", stepSw.ElapsedMilliseconds);

        _logger.LogInformation("[PERF] 좌석 선택 전체 소요: {Ms}ms", totalSw.ElapsedMilliseconds);
    }

    private async Task<IFrame> FindMelonSeatFrameAsync(IPage page, TimeSpan timeout, CancellationToken cancellationToken)
    {
        foreach (var frame in page.Frames)
        {
            if (frame == page.MainFrame) continue;
            if (!frame.Url.Contains("stepSeat.htm") && !frame.Url.Contains("stepBlock.htm")) continue;
            try
            {
                var canvas = frame.Locator("#ez_canvas");
                if (await canvas.CountAsync() > 0)
                {
                    _logger.LogInformation("[FindSeatFrame] fast-path: #ez_canvas 즉시 발견. frameUrl={Url}", frame.Url);
                    return frame;
                }
            }
            catch (PlaywrightException) { }
        }

        var deadline = DateTimeOffset.UtcNow + timeout;
        var loggedOnce = false;
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (!loggedOnce)
            {
                var frameUrls = page.Frames.Select(f => f.Url).ToList();
                _logger.LogInformation("[FindSeatFrame] fast-path 실패, 폴링 시작. pageUrl={Url}, frameCount={Count}, frameUrls={FrameUrls}",
                    SafePageUrl(page), frameUrls.Count, string.Join(" | ", frameUrls));
                loggedOnce = true;
            }

            foreach (var frame in page.Frames)
            {
                if (frame == page.MainFrame) continue;
                try
                {
                    var canvas = frame.Locator("#ez_canvas");
                    if (await canvas.CountAsync() > 0)
                    {
                        _logger.LogInformation("[FindSeatFrame] #ez_canvas 발견! frameUrl={Url}", frame.Url);
                        return frame;
                    }
                }
                catch (PlaywrightException) { }
            }

            await Task.Delay(PollDelayMilliseconds, cancellationToken);
        }

        var finalFrameUrls = page.Frames.Select(f => f.Url).ToList();
        _logger.LogError("[FindSeatFrame] 타임아웃! pageUrl={Url}, frameCount={Count}, frameUrls={FrameUrls}",
            SafePageUrl(page), finalFrameUrls.Count, string.Join(" | ", finalFrameUrls));
        throw new TimeoutException("멜론 좌석맵 iframe(#ez_canvas)을 찾지 못했습니다.");
    }

    private static async Task<bool> IsMelonZoneSelectionRequiredAsync(IFrame seatFrame)
    {
        try
        {
            var zoneCanvas = seatFrame.Locator("#ez_canvas_zone");
            if (await zoneCanvas.CountAsync() > 0)
                return true;

            var seatInfo = seatFrame.Locator("#txtSelectSeatInfo");
            if (await seatInfo.CountAsync() > 0)
            {
                var text = await seatInfo.InnerTextAsync();
                if (text.Contains("구역을 먼저", StringComparison.OrdinalIgnoreCase))
                    return true;
            }
        }
        catch (PlaywrightException) { }

        return false;
    }

    private async Task WaitForMelonZoneSelectionAsync(IFrame seatFrame, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        progress?.Report(new AutomationProgress("구역 선택 대기 중", "구역 선택 대기 — 사용자 클릭 필요"));

        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var rectCount = await seatFrame.Locator("#ez_canvas rect").CountAsync();
                if (rectCount > 50)
                {
                    var smallRects = await seatFrame.EvaluateAsync<int>(@"() => {
                    const rects = document.querySelectorAll('#ez_canvas rect');
                    let count = 0;
                    for (const r of rects) {
                        const w = parseFloat(r.getAttribute('width'));
                        if (w > 0 && w <= 15) count++;
                    }
                    return count;
                }");
                    if (smallRects > 10)
                    {
                        _logger.LogInformation("구역 선택 후 좌석 {Count}개 로드 감지.", smallRects);
                        return;
                    }
                }
            }
            catch (PlaywrightException) { }

            await Task.Delay(50, cancellationToken);
        }
    }

    private async Task<IFrame> SelectMelonSeatInFrameAsync(IPage page, IFrame seatFrame, TimeSpan timeout, CancellationToken cancellationToken)
    {
        const int maxRetries = 10;
        var currentFrame = seatFrame;

        for (var retry = 0; retry < maxRetries; retry++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var scanClickResult = await currentFrame.EvaluateAsync<string>(@"(args) => {
                const rects = document.querySelectorAll('#ez_canvas rect');
                const seats = [];
                let i = 0;
                for (const r of rects) {
                    const fill = r.getAttribute('fill') || 'none';
                    const w = parseFloat(r.getAttribute('width'));
                    const h = parseFloat(r.getAttribute('height'));
                    if (w > 0 && w <= 15 && h > 0 && h <= 15 && fill !== 'none' && fill.toUpperCase() !== '#DDDDDD') {
                        seats.push({ x: parseFloat(r.getAttribute('x')), y: parseFloat(r.getAttribute('y')), idx: i });
                    }
                    i++;
                }
                if (seats.length === 0) return JSON.stringify({ s: 'empty', c: 0 });
                seats.sort((a, b) => a.y - b.y || a.x - b.x);
                const ti = seats.length >= args.offset ? args.offset - 1 : 0;
                const t = seats[ti];
                const rect = rects[t.idx];
                if (!rect) return JSON.stringify({ s: 'not_found', c: seats.length });
                const evt = document.createEvent('MouseEvents');
                evt.initMouseEvent('click', true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
                rect.dispatchEvent(evt);
                return JSON.stringify({ s: 'clicked', c: seats.length, ti: ti, x: t.x, y: t.y });
            }", new { offset = SeatSelectionOffset });

                using var scanDoc = JsonDocument.Parse(scanClickResult);
                var status = scanDoc.RootElement.GetProperty("s").GetString();
                var seatCount = scanDoc.RootElement.GetProperty("c").GetInt32();

                if (status == "empty")
                {
                    _logger.LogWarning("선택 가능한 좌석 없음. retry={Retry}", retry);
                    await Task.Delay(50, cancellationToken);
                    continue;
                }

                if (status != "clicked")
                {
                    _logger.LogWarning("좌석 클릭 실패: status={Status}, count={Count}", status, seatCount);
                    continue;
                }

                _logger.LogInformation("좌석 스캔+클릭 완료: available={Count}, targetIdx={Idx}, x={X}, y={Y}",
                    seatCount, scanDoc.RootElement.GetProperty("ti").GetInt32(),
                    scanDoc.RootElement.GetProperty("x").GetDouble(), scanDoc.RootElement.GetProperty("y").GetDouble());

                var validation = await currentFrame.EvaluateAsync<string>(@"() => {
                const ad = window.__melonAlertDetected === true;
                const sel = document.querySelectorAll('#partSeatSelected li').length;
                return JSON.stringify({ ad: ad, sel: sel });
            }");

                using var valDoc = JsonDocument.Parse(validation);
                var alertDetected = valDoc.RootElement.GetProperty("ad").GetBoolean();
                var selectedCount = valDoc.RootElement.GetProperty("sel").GetInt32();

                if (alertDetected)
                {
                    _logger.LogInformation("좌석 중복 선택 감지 — 다른 좌석으로 재시도. retry={Retry}", retry);
                    await DismissMelonSeatConflictAlertAsync(currentFrame);
                    continue;
                }

                if (selectedCount > 0)
                {
                    _logger.LogInformation("좌석 선택 성공. selectedCount={Count}", selectedCount);
                    return currentFrame;
                }

                _logger.LogWarning("좌석 선택 후 선택된 좌석 목록 비어있음. retry={Retry}", retry);
            }
            catch (PlaywrightException ex)
            {
                _logger.LogWarning("좌석 선택 중 frame detached 감지 — frame 재탐색. retry={Retry}, error={Error}", retry, ex.Message);
                await Task.Delay(100, cancellationToken);
                currentFrame = await FindMelonSeatFrameAsync(page, timeout, cancellationToken);
            }
        }

        throw new InvalidOperationException($"좌석 선택 실패 ({maxRetries}회 시도).");
    }

    private async Task<bool> DetectMelonSeatConflictAsync(IFrame seatFrame)
    {
        try
        {
            var page = seatFrame.Page;
            var dialogDetected = await page.EvaluateAsync<bool>(@"() => {
            return window.__melonAlertDetected === true;
        }");
            return dialogDetected;
        }
        catch (PlaywrightException)
        {
            return false;
        }
    }

    private async Task DismissMelonSeatConflictAlertAsync(IFrame seatFrame)
    {
        try
        {
            var page = seatFrame.Page;
            await page.EvaluateAsync(@"() => { window.__melonAlertDetected = false; }");
        }
        catch (PlaywrightException) { }
    }

    private async Task ClickMelonSeatCompleteAsync(IPage page, IFrame seatFrame, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var currentFrame = seatFrame;
        const int maxFrameRetries = 3;

        for (var attempt = 0; attempt < maxFrameRetries; attempt++)
        {
            try
            {
                var nextBtn = currentFrame.Locator("#nextTicketSelection");
                await WaitForConditionAsync(
                    async () => await nextBtn.CountAsync() > 0,
                    timeout,
                    cancellationToken,
                    "멜론 '좌석 선택 완료' 버튼을 찾지 못했습니다.");

                try
                {
                    await nextBtn.First.EvaluateAsync("el => el.click()");
                }
                catch (PlaywrightException)
                {
                    await nextBtn.First.ClickAsync(new LocatorClickOptions { Timeout = 1000, Force = true });
                }

                await Task.Delay(200, cancellationToken);

                if (await DetectMelonSeatConflictAsync(currentFrame))
                {
                    _logger.LogWarning("좌석 선택 완료 클릭 후 중복 좌석 감지 — 좌석 재선택 필요. attempt={Attempt}", attempt);
                    await DismissMelonSeatConflictAlertAsync(currentFrame);
                    throw new InvalidOperationException("좌석 선택 완료 시 중복 좌석 감지됨 — 재시도 필요.");
                }

                _logger.LogInformation("멜론 '좌석 선택 완료' 버튼 클릭 완료.");
                return;
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (PlaywrightException ex) when (attempt < maxFrameRetries - 1)
            {
                _logger.LogWarning("좌석 선택 완료 버튼 클릭 중 frame detached 감지 — frame 재탐색. attempt={Attempt}, error={Error}", attempt, ex.Message);
                await Task.Delay(100, cancellationToken);
                currentFrame = await FindMelonSeatFrameAsync(page, timeout, cancellationToken);
            }
        }
    }
    private static async Task EnsureMelonPopupClosedAsync(IPage page, TimeSpan timeout, CancellationToken cancellationToken)
    {
        try
        {
            var hasPopup = await page.EvaluateAsync<bool>("""
                () => {
                    const candidates = Array.from(document.querySelectorAll('#popup_notice, .popup, [class*="popup"]'));
                    return candidates.some(element => {
                        const style = window.getComputedStyle(element);
                        return style.display !== 'none' && style.visibility !== 'hidden' && Number(style.opacity || '1') > 0;
                    });
                }
                """);
            if (!hasPopup) return;
        }
        catch (PlaywrightException) { return; }

        var deadline = DateTimeOffset.UtcNow + timeout;
        var selectors = new[]
        {
            "a:has-text('레이어팝업닫기')",
            "a.btn_layerpopup_close",
            "#popup_notice .close",
            "#popup_notice [class*='close']",
            ".popup .close",
            ".popup .btn_close",
            "[class*='popup'] [class*='close']",
            "[class*='Close']"
        };

        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (page.IsClosed)
                return;

            try
            {
                foreach (var selector in selectors)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var buttons = page.Locator(selector);
                    var count = await buttons.CountAsync();
                    for (var index = 0; index < count; index++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        var button = buttons.Nth(index);
                        if (!await button.IsVisibleAsync())
                        {
                            continue;
                        }

                        try
                        {
                            await button.ClickAsync(new LocatorClickOptions
                            {
                                Force = true,
                                Timeout = 300
                            });
                        }
                        catch (PlaywrightException)
                        {
                            try
                            {
                                await button.EvaluateAsync("element => { element.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window })); if (typeof element.click === 'function') { element.click(); } }");
                            }
                            catch (PlaywrightException)
                            {
                            }
                        }

                        await Task.Delay(50, cancellationToken);
                        break;
                    }
                }

                cancellationToken.ThrowIfCancellationRequested();

                var hasVisiblePopup = await page.EvaluateAsync<bool>("""
                    () => {
                        const candidates = Array.from(document.querySelectorAll('#popup_notice, .popup, [class*="popup"]'));
                        return candidates.some(element => {
                            const style = window.getComputedStyle(element);
                            return style.display !== 'none' && style.visibility !== 'hidden' && Number(style.opacity || '1') > 0;
                        });
                    }
                    """);

                if (!hasVisiblePopup)
                {
                    return;
                }
            }
            catch (PlaywrightException) when (page.IsClosed)
            {
                return;
            }
            catch (PlaywrightException)
            {
            }

            await Task.Delay(PollDelayMilliseconds, cancellationToken);
        }
    }
    private static async Task SelectMelonDateAsync(IPage page, DateOnly desiredDate, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var desiredPerfday = desiredDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
        var listItems = page.Locator("#list_date .item_date[data-perfday]");
        var calendarButton = page.Locator($"#cal_wrapper .ticketCalendarBtn[data-perfday='{desiredPerfday}']:not([disabled])").First;

        await WaitForConditionAsync(
            async () => await listItems.CountAsync() > 0 || await calendarButton.CountAsync() > 0,
            timeout,
            cancellationToken,
            "Melon 날짜 목록을 찾지 못했습니다.");

        var listCount = await listItems.CountAsync();
        for (var index = 0; index < listCount; index++)
        {
            var item = listItems.Nth(index);
            if (!string.Equals(await item.GetAttributeAsync("data-perfday"), desiredPerfday, StringComparison.Ordinal))
            {
                continue;
            }

            await ClickNolElementAsync(item);
            await WaitForConditionAsync(
                async () => (await item.GetAttributeAsync("class") ?? string.Empty).Contains("on", StringComparison.OrdinalIgnoreCase),
                timeout,
                cancellationToken,
                "Melon 관람일 선택 반영을 확인하지 못했습니다.");
            return;
        }

        if (await calendarButton.CountAsync() > 0)
        {
            await ClickNolElementAsync(calendarButton);
            await WaitForConditionAsync(
                async () => (await calendarButton.GetAttributeAsync("class") ?? string.Empty).Contains("on", StringComparison.OrdinalIgnoreCase) ||
                          await page.Locator($"#list_date .item_date[data-perfday='{desiredPerfday}'].on").CountAsync() > 0,
                timeout,
                cancellationToken,
                "Melon 관람일 선택 반영을 확인하지 못했습니다.");
            return;
        }

        throw new InvalidOperationException($"Melon 관람일을 찾지 못했습니다: {desiredDate:yyyy.MM.dd}");
    }
    private static async Task SelectMelonTimeAsync(IPage page, string desiredTime, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var timesLocator = page.Locator("#list_time .item_time");

        await WaitForConditionAsync(async () => await timesLocator.CountAsync() > 0, timeout, cancellationToken, "Melon 시간 목록을 찾지 못했습니다.");

        var count = await timesLocator.CountAsync();
        for (var index = 0; index < count; index++)
        {
            var item = timesLocator.Nth(index);
            var className = await item.GetAttributeAsync("class") ?? string.Empty;
            if (className.Contains("disabled", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var timeText = NormalizeText(await item.InnerTextAsync());
            if (!IsMatchingMelonTime(timeText, desiredTime))
            {
                continue;
            }

            if (className.Contains("on", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            await ClickNolElementAsync(item);
            await WaitForConditionAsync(
                async () => (await item.GetAttributeAsync("class") ?? string.Empty).Contains("on", StringComparison.OrdinalIgnoreCase),
                timeout,
                cancellationToken,
                "Melon 시간 선택 반영을 확인하지 못했습니다.");
            return;
        }

        throw new InvalidOperationException($"Melon 시간을 찾지 못했습니다: {desiredTime}");
    }
    private async Task<IPage> ClickMelonBookingAsync(IPage page, TimeSpan timeout, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var existingPopup = page.Context.Pages
            .FirstOrDefault(p => p != page && !p.IsClosed && SafePageUrl(p).Contains("/reservation/popup/onestop.htm", StringComparison.OrdinalIgnoreCase));
        if (existingPopup is not null)
        {
            _logger.LogInformation("[Melon] 기존 onestop.htm 팝업 재사용. url={Url}", SafePageUrl(existingPopup));
            await existingPopup.BringToFrontAsync();
            return existingPopup;
        }

        var bookingButton = page.Locator("#ticketReservation_Btn").First;
        if (await bookingButton.CountAsync() == 0)
        {
            throw new InvalidOperationException("Melon 예매 버튼을 찾지 못했습니다.");
        }

        var beforePages = page.Context.Pages.Where(x => !x.IsClosed).ToHashSet();
        await bookingButton.ScrollIntoViewIfNeededAsync();
        _logger.LogInformation("예매하기 버튼 클릭 시도. beforePagesCount={Count}", beforePages.Count);
        await ClickNolBookingButtonAsync(bookingButton, timeout);

        var deadline = DateTimeOffset.UtcNow + timeout;
        IPage? queuePage = null;

        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var openPages = page.Context.Pages.Where(x => !x.IsClosed).ToList();
            var newPage = openPages.FirstOrDefault(x => !beforePages.Contains(x));
            if (newPage is not null)
            {
                if (SafePageUrl(newPage).Contains("/reservation/popup/onestop.htm", StringComparison.OrdinalIgnoreCase))
                {
                    await newPage.BringToFrontAsync();
                    return newPage;
                }

                queuePage = newPage;
            }

            var foundPopup = openPages.FirstOrDefault(x => SafePageUrl(x).Contains("/reservation/popup/onestop.htm", StringComparison.OrdinalIgnoreCase));
            if (foundPopup is not null)
            {
                await foundPopup.BringToFrontAsync();
                return foundPopup;
            }

            await Task.Delay(PollDelayMilliseconds, cancellationToken);
        }

        if (queuePage is not null && !queuePage.IsClosed)
        {
            _logger.LogInformation("[Melon] 대기열 감지 — onestop.htm 전환까지 무한 대기. queueUrl={Url}", SafePageUrl(queuePage));
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

                if (queuePage.IsClosed)
                    throw new InvalidOperationException("Melon 대기열 페이지가 닫혔습니다.");

                if (SafePageUrl(queuePage).Contains("/reservation/popup/onestop.htm", StringComparison.OrdinalIgnoreCase))
                {
                    await queuePage.BringToFrontAsync();
                    _logger.LogInformation("[Melon] 대기열 종료 — onestop.htm 도착. url={Url}", SafePageUrl(queuePage));
                    return queuePage;
                }

                var allPages = page.Context.Pages.Where(x => !x.IsClosed).ToList();
                var popup = allPages.FirstOrDefault(x => SafePageUrl(x).Contains("/reservation/popup/onestop.htm", StringComparison.OrdinalIgnoreCase));
                if (popup is not null)
                {
                    await popup.BringToFrontAsync();
                    _logger.LogInformation("[Melon] 대기열 종료 — 새 onestop.htm 팝업 감지. url={Url}", SafePageUrl(popup));
                    return popup;
                }

                await Task.Delay(100, cancellationToken);
            }
        }

        throw new TimeoutException("Melon 예매 버튼 클릭 후 팝업 전환을 확인하지 못했습니다.");
    }
    private async Task SolveMelonCaptchaAsync(IPage page, TimeSpan timeout, CancellationToken cancellationToken)
    {
        const int maxAttempts = 5;
        const int melonCaptchaLength = 6;

        _logger.LogInformation("CAPTCHA 입력창 대기 시작 (최대 500ms). url={Url}", SafePageUrl(page));
        var (inputLocator, captchaFrame) = await FindMelonCaptchaInputAsync(page, TimeSpan.FromMilliseconds(500), cancellationToken);
        if (inputLocator is null)
        {
            foreach (var contextPage in page.Context.Pages.Where(p => p != page && !p.IsClosed))
            {
                (inputLocator, captchaFrame) = await FindMelonCaptchaInputAsync(contextPage, TimeSpan.FromMilliseconds(300), cancellationToken);
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
                text = await RecognizeCaptchaTextAsync(inputLocator, page, captchaFrame, cancellationToken);
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
                    await TryRefreshMelonCaptchaImageAsync(page, captchaFrame, cancellationToken);
                continue;
            }

            if (string.IsNullOrEmpty(text))
            {
                _logger.LogWarning("[CAPTCHA] OCR 빈 결과 (attempt={Attempt}, {Ms}ms) — 이미지 미로드 또는 인식 실패, 재시도.", attempt, attemptSw.ElapsedMilliseconds);
                if (attempt < maxAttempts)
                    await TryRefreshMelonCaptchaImageAsync(page, captchaFrame, cancellationToken);
                continue;
            }

            _logger.LogInformation("CAPTCHA attempt {Attempt}/{Max}: text={Text} ocrMs={OcrMs}", attempt, maxAttempts, text, attemptSw.ElapsedMilliseconds);

            if (text.Length != melonCaptchaLength)
            {
                if (attempt < maxAttempts)
                    await TryRefreshMelonCaptchaImageAsync(page, captchaFrame, cancellationToken);
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
                    await TryRefreshMelonCaptchaImageAsync(page, captchaFrame, cancellationToken);
                continue;
            }

            _logger.LogInformation("[CAPTCHA] CAPTCHA 제출 완료 (alert 없음). 좌석 진행. attempt={Attempt}, totalMs={Ms}", attempt, attemptSw.ElapsedMilliseconds);
            return;
        }

        _logger.LogWarning("[CAPTCHA] CAPTCHA 자동 인식 {Max}회 모두 실패 — 좌석 선택 진행 시도.", maxAttempts);
    }

    private async Task TryRefreshMelonCaptchaImageAsync(IPage page, IFrame? captchaFrame, CancellationToken cancellationToken)
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

    private static async Task<(ILocator? inputLocator, IFrame? frame)> FindMelonCaptchaInputAsync(
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

            await Task.Delay(PollDelayMilliseconds, cancellationToken);
        }

        return (null, null);
    }

    private static bool IsMatchingMelonTime(string actual, string desired)
    {
        var actualDigits = DigitsOnlyPattern.Replace(actual, string.Empty);
        var desiredDigits = DigitsOnlyPattern.Replace(desired, string.Empty);
        if (actualDigits.Length >= 4 && desiredDigits.Length >= 4)
        {
            return actualDigits[..4] == desiredDigits[..4];
        }

        return string.Equals(NormalizeText(actual), NormalizeText(desired), StringComparison.Ordinal);
    }
    private sealed record MelonSeatInfo(
        [property: System.Text.Json.Serialization.JsonPropertyName("x")] double X,
        [property: System.Text.Json.Serialization.JsonPropertyName("y")] double Y,
        [property: System.Text.Json.Serialization.JsonPropertyName("idx")] int Idx);
}
