using MockTicketServer.Pages;

var queueSeconds = args.Length > 0 && int.TryParse(args[0], out var s) ? s : 60;
var hasCaptcha = !args.Contains("--no-captcha", StringComparer.OrdinalIgnoreCase);
var hasZone = !args.Contains("--no-zone", StringComparer.OrdinalIgnoreCase);
var conflictSeats = args
    .Select(arg => arg.Split('=', 2))
    .Where(parts => parts.Length == 2 && parts[0].Equals("--conflict-seats", StringComparison.OrdinalIgnoreCase))
    .Select(parts => int.TryParse(parts[1], out var value) ? Math.Max(0, value) : 0)
    .FirstOrDefault();

var filteredArgs = args.Where(a =>
    !a.StartsWith("--no-", StringComparison.OrdinalIgnoreCase) &&
    !a.StartsWith("--conflict-seats=", StringComparison.OrdinalIgnoreCase)).ToArray();
var builder = WebApplication.CreateBuilder(filteredArgs);
var port = Environment.GetEnvironmentVariable("MOCK_PORT") ?? "8080";
builder.WebHost.UseUrls($"http://0.0.0.0:{port}");
builder.Logging.SetMinimumLevel(LogLevel.Information);

var app = builder.Build();

app.Logger.LogInformation("Mock Ticket Server 시작. 대기열: {Queue}초, 캡차: {Captcha}, 구역: {Zone}, 충돌 좌석: {ConflictSeats}, 포트: {Port}",
    queueSeconds, hasCaptcha, hasZone, conflictSeats, port);

app.Use(async (context, next) =>
{
    var host = context.Request.Host.Host;
    context.Items["SiteType"] = host switch
    {
        var h when h.Contains("interpark", StringComparison.OrdinalIgnoreCase) => "nol",
        var h when h.Contains("melon", StringComparison.OrdinalIgnoreCase) => "melon",
        _ => "unknown"
    };
    await next();
});

// ──────────────── NOL Routes ────────────────

app.MapGet("/goods/{id}", (string id, HttpContext ctx) =>
{
    if ((string)ctx.Items["SiteType"]! != "nol")
        return Results.NotFound("NOL 전용 경로입니다.");
    app.Logger.LogInformation("[NOL] 상품 페이지 요청. id={Id}", id);
    return Results.Content(NolPages.GoodsPage(queueSeconds), "text/html; charset=utf-8");
});

app.MapGet("/queue", (HttpContext ctx) =>
{
    if ((string)ctx.Items["SiteType"]! != "nol")
        return Results.NotFound("NOL 전용 경로입니다.");
    app.Logger.LogInformation("[NOL] 대기열 페이지 요청. duration={Seconds}s", queueSeconds);
    return Results.Content(NolPages.QueuePage(queueSeconds), "text/html; charset=utf-8");
});

app.MapGet("/captcha", (HttpContext ctx) =>
{
    if ((string)ctx.Items["SiteType"]! != "nol")
        return Results.NotFound("NOL 전용 경로입니다.");
    app.Logger.LogInformation("[NOL] 캡차 페이지 요청.");
    return Results.Content(NolPages.CaptchaPage(hasCaptcha), "text/html; charset=utf-8");
});

// ──────────────── Melon Routes ────────────────

app.MapGet("/performance/index.htm", (HttpContext ctx) =>
{
    if ((string)ctx.Items["SiteType"]! != "melon")
        return Results.NotFound("Melon 전용 경로입니다.");
    app.Logger.LogInformation("[Melon] 공연 페이지 요청.");
    return Results.Content(MelonPages.PerformancePage(queueSeconds), "text/html; charset=utf-8");
});

app.MapGet("/queue/popup", (HttpContext ctx) =>
{
    if ((string)ctx.Items["SiteType"]! != "melon")
        return Results.NotFound("Melon 전용 경로입니다.");
    app.Logger.LogInformation("[Melon] 대기열 팝업 요청. duration={Seconds}s", queueSeconds);
    return Results.Content(MelonPages.QueuePopup(queueSeconds), "text/html; charset=utf-8");
});

app.MapGet("/reservation/popup/onestop.htm", (HttpContext ctx) =>
{
    if ((string)ctx.Items["SiteType"]! != "melon")
        return Results.NotFound("Melon 전용 경로입니다.");
    app.Logger.LogInformation("[Melon] 예매 팝업(onestop) 요청.");
    return Results.Content(MelonPages.OnestopPopup(hasCaptcha), "text/html; charset=utf-8");
});

app.MapGet("/reservation/popup/stepSeat.htm", (HttpContext ctx) =>
{
    if ((string)ctx.Items["SiteType"]! != "melon")
        return Results.NotFound("Melon 전용 경로입니다.");
    app.Logger.LogInformation("[Melon] 좌석 프레임 요청.");
    return Results.Content(MelonPages.SeatFrame(hasZone, conflictSeats), "text/html; charset=utf-8");
});

// ──────────────── Fallback ────────────────

app.MapGet("/", (HttpContext ctx) =>
{
    var siteType = (string)ctx.Items["SiteType"]!;
    return siteType switch
    {
        "nol" => Results.Redirect("/goods/12345"),
        "melon" => Results.Redirect("/performance/index.htm"),
        _ => Results.Content("""
            <html><body style="font-family:sans-serif;text-align:center;padding:50px;">
            <h1>Mock Ticket Server</h1>
            <p>Host 헤더로 사이트를 구분합니다.</p>
            <ul style="list-style:none;">
            <li><a href="http://tickets.interpark.com/goods/12345">NOL (인터파크)</a></li>
            <li><a href="http://ticket.melon.com/performance/index.htm">Melon (멜론)</a></li>
            </ul>
            </body></html>
            """, "text/html; charset=utf-8")
    };
});

app.Run();
