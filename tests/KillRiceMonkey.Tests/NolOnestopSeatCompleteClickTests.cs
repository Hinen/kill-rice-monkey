using System.Net;
using Microsoft.Playwright;

namespace KillRiceMonkey.Tests;

/// <summary>
/// NOL onestop "선택 완료" 버튼 클릭 회귀 테스트.
/// 실제 버그: JS el.click() (untrusted) 은 React EntButton 의 onClick 을 발동시키지 못하여
/// 프로그램이 "성공 종료"로 판단했으나 실제 페이지 전환이 일어나지 않음.
/// 수정: Playwright ClickAsync (CDP trusted click) 사용 + 페이지 전환 검증 + 실패 시 throw.
/// </summary>
[Collection("Playwright")]
public class NolOnestopSeatCompleteClickTests : IAsyncLifetime
{
    private HttpListener? _server;
    private IPlaywright? _playwright;
    private IBrowser? _browser;
    private string _baseUrl = string.Empty;

    public async Task InitializeAsync()
    {
        // 동적 포트 할당: 빈 포트를 찾아서 사용
        int port;
        using (var tempSocket = new System.Net.Sockets.TcpListener(IPAddress.Loopback, 0))
        {
            tempSocket.Start();
            port = ((IPEndPoint)tempSocket.LocalEndpoint).Port;
            tempSocket.Stop();
        }

        _baseUrl = $"http://127.0.0.1:{port}";
        _server = new HttpListener();
        _server.Prefixes.Add($"{_baseUrl}/");
        _server.Start();

        _ = Task.Run(async () =>
        {
            while (_server.IsListening)
            {
                HttpListenerContext? ctx;
                try { ctx = await _server.GetContextAsync(); }
                catch { break; }

                var path = ctx.Request.Url?.AbsolutePath ?? "/";
                string html;
                if (path.StartsWith("/onestop/seat"))
                    html = MockOnestopSeatHtml();
                else if (path.StartsWith("/onestop/complete"))
                    html = "<html><body><h1>COMPLETE</h1></body></html>";
                else
                    html = "<html><body>404</body></html>";

                var buffer = System.Text.Encoding.UTF8.GetBytes(html);
                ctx.Response.ContentType = "text/html; charset=utf-8";
                ctx.Response.ContentLength64 = buffer.Length;
                await ctx.Response.OutputStream.WriteAsync(buffer);
                ctx.Response.Close();
            }
        });

        _playwright = await Playwright.CreateAsync();
        _browser = await _playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions { Headless = true });
    }

    public async Task DisposeAsync()
    {
        if (_browser is not null) await _browser.DisposeAsync();
        _playwright?.Dispose();
        _server?.Stop();
        _server?.Close();
    }

    /// <summary>
    /// Playwright ClickAsync (trusted click) 은 "선택 완료" 버튼을 정상적으로 클릭하여
    /// /onestop/complete 로 이동해야 한다.
    /// </summary>
    [Fact]
    public async Task TrustedClick_NavigatesToCompletePage()
    {
        var page = await _browser!.NewPageAsync();
        await page.GotoAsync($"{_baseUrl}/onestop/seat");

        // 좌석 선택
        var seat = page.Locator("[class*='SeatMap_seatGroup'] circle:not(.disabled)").First;
        await seat.ClickAsync();

        // "선택 완료" 버튼이 활성화됨을 확인
        var completeBtn = page.Locator("button:has-text('선택 완료')").First;
        await Assertions.Expect(completeBtn).ToBeEnabledAsync();

        // Playwright trusted click → 페이지 전환 발생
        await completeBtn.ClickAsync();
        await page.WaitForURLAsync("**/onestop/complete**", new PageWaitForURLOptions { Timeout = 3000 });

        Assert.Contains("/onestop/complete", page.Url);
        await page.CloseAsync();
    }

    /// <summary>
    /// JS el.click() (untrusted click) 은 "선택 완료" 버튼의 isTrusted 검사에 의해 무시되어
    /// 페이지 전환이 일어나지 않아야 한다. 이것이 원래 버그의 재현이다.
    /// </summary>
    [Fact]
    public async Task UntrustedJsClick_DoesNotNavigate()
    {
        var page = await _browser!.NewPageAsync();
        await page.GotoAsync($"{_baseUrl}/onestop/seat");

        // 좌석 선택
        var seat = page.Locator("[class*='SeatMap_seatGroup'] circle:not(.disabled)").First;
        await seat.ClickAsync();

        // "선택 완료" 버튼 활성화 확인
        var completeBtn = page.Locator("button:has-text('선택 완료')").First;
        await Assertions.Expect(completeBtn).ToBeEnabledAsync();

        // JS el.click() (untrusted) — 이것이 원래 버그의 원인
        await completeBtn.EvaluateAsync("el => el.click()");
        await Task.Delay(500);

        // URL이 변경되지 않아야 한다 (untrusted click 은 무시됨)
        Assert.Contains("/onestop/seat", page.Url);
        Assert.DoesNotContain("/onestop/complete", page.Url);
        await page.CloseAsync();
    }

    private static string MockOnestopSeatHtml() => """
<!DOCTYPE html>
<html lang="ko">
<head><meta charset="utf-8" /><title>Mock onestop/seat</title></head>
<body>
<div id="__next">
    <div class="SeatPlan_seatPlan__mock">
        <div class="SeatMap_seatGroup__mock">
            <svg viewBox="0 0 200 100">
                <circle cx="50" cy="50" r="8" fill="#6a6" />
                <circle cx="80" cy="50" r="8" fill="#6a6" />
                <circle cx="110" cy="50" r="8" fill="#6a6" class="disabled" />
            </svg>
        </div>
    </div>
    <div class="InfoSelected_outerWrap__mock" id="selectedInfo">선택한 좌석이 없습니다.</div>
    <button class="EntButton_button__mock EntButton_primary__mock" id="completeBtn" disabled>선택 완료</button>
</div>
<script>
    var selectedSeat = null;
    var circles = document.querySelectorAll('[class*="SeatMap_seatGroup"] circle:not(.disabled)');
    var info = document.getElementById('selectedInfo');
    var completeBtn = document.getElementById('completeBtn');

    circles.forEach(function(circle) {
        circle.addEventListener('click', function() {
            if (selectedSeat) selectedSeat.classList.remove('selected');
            circle.classList.add('selected');
            selectedSeat = circle;
            info.textContent = '좌석 선택됨';
            completeBtn.disabled = false;
        });
    });

    /* isTrusted=false 이벤트 무시 — 실제 NOL 사이트의 EntButton 동작 시뮬레이션 */
    completeBtn.addEventListener('click', function(e) {
        if (!e.isTrusted) return;
        if (completeBtn.disabled) return;
        window.location.href = '/onestop/complete';
    });
</script>
</body>
</html>
""";
}
