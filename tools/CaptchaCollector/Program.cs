using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Playwright;

const string NolCdpEndpoint = "http://localhost:9222";
const string MelonCdpEndpoint = "http://localhost:9223";
const string ImageSelector = "#imgCaptcha, #captchaImg, [class*='captchaImage'] img, img[src*='captcha' i], img[src*='cap_img' i]";
const string RefreshSelector = "#divRecaptcha .capchaBtns a:last-of-type, #btnReload, .refreshBtn, [class*='buttonRefresh'], button[aria-label*='새 문자']";

var targetCount = args.Length > 0 && int.TryParse(args[0], out var c) ? c : 1000;
var captchaType = args.Length > 1 ? args[1].ToLowerInvariant() : "new";
var isMelon = captchaType == "melon";
var cdpEndpoint = isMelon ? MelonCdpEndpoint : NolCdpEndpoint;
var apiKey = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY");

var projectDir = AppDomain.CurrentDomain.BaseDirectory;
var repoRoot = Path.GetFullPath(Path.Combine(projectDir, "..", "..", "..", "..", ".."));
var samplesDir = Path.Combine(repoRoot, "captcha-samples", captchaType);
Directory.CreateDirectory(samplesDir);

var groundTruthPath = Path.Combine(samplesDir, "ground_truth.csv");
var existingLabels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
if (File.Exists(groundTruthPath))
{
    foreach (var line in await File.ReadAllLinesAsync(groundTruthPath))
    {
        var parts = line.Split(',', 2);
        if (parts.Length == 2 && parts[1].Length == 6)
            existingLabels[parts[0]] = parts[1];
    }
}

var existingCount = Directory.GetFiles(samplesDir, "captcha_*.png").Length;
Console.WriteLine($"CAPTCHA Collector: type={captchaType.ToUpperInvariant()}, target={targetCount}, existing={existingCount}");
Console.WriteLine($"Output: {samplesDir}");
Console.WriteLine($"Vision API: {(string.IsNullOrWhiteSpace(apiKey) ? "OFF (no ANTHROPIC_API_KEY)" : "ON")}");
Console.WriteLine();

if (existingCount >= targetCount)
{
    Console.WriteLine($"Already have {existingCount} samples (>= {targetCount}). Nothing to do.");
    return 0;
}

using var playwright = await Playwright.CreateAsync();
IBrowser browser;
try
{
    browser = await playwright.Chromium.ConnectOverCDPAsync(cdpEndpoint);
}
catch (Exception ex)
{
    Console.Error.WriteLine($"CDP 연결 실패: {ex.Message}");
    var buttonName = isMelon ? "Melon Remote Debug 열기" : "NOL Remote Debug 열기";
    Console.Error.WriteLine($"remote-debug 브라우저를 먼저 실행하세요 (앱에서 '{buttonName}' 버튼).");
    return 1;
}

Console.WriteLine($"CDP 연결 성공. contexts={browser.Contexts.Count}");

var context = browser.Contexts.FirstOrDefault();
if (context is null)
{
    Console.Error.WriteLine("브라우저 컨텍스트를 찾을 수 없습니다.");
    return 1;
}

IPage? captchaPage = null;
ILocator? imageLocator = null;
IFrame? captchaFrame = null;

foreach (var pg in context.Pages)
{
    var imgLoc = pg.Locator(ImageSelector);
    try
    {
        if (await imgLoc.CountAsync() > 0)
        {
            captchaPage = pg;
            imageLocator = imgLoc;
            captchaFrame = null;
            break;
        }
    }
    catch { }

    foreach (var frame in pg.Frames)
    {
        if (frame == pg.MainFrame) continue;
        var frameImg = frame.Locator(ImageSelector);
        try
        {
            if (await frameImg.CountAsync() > 0)
            {
                captchaPage = pg;
                imageLocator = frameImg;
                captchaFrame = frame;
                break;
            }
        }
        catch { }
    }

    if (imageLocator is not null) break;
}

if (captchaPage is null || imageLocator is null)
{
    Console.Error.WriteLine("CAPTCHA 이미지를 찾을 수 없습니다.");
    Console.Error.WriteLine($"탐색한 페이지: {context.Pages.Count}개");
    foreach (var pg in context.Pages)
    {
        Console.Error.WriteLine($"  - {pg.Url}");
        foreach (var frame in pg.Frames)
        {
            if (frame == pg.MainFrame) continue;
            Console.Error.WriteLine($"    frame: {frame.Url}");
        }
    }
    Console.Error.WriteLine($"사용 셀렉터: {ImageSelector}");
    Console.Error.WriteLine("CAPTCHA가 표시된 페이지에서 다시 시도하세요.");
    return 1;
}

var frameInfo = captchaFrame is not null ? $" (frame: {captchaFrame.Url[..Math.Min(60, captchaFrame.Url.Length)]}...)" : "";
Console.WriteLine($"CAPTCHA 이미지 발견. url={SafeUrl(captchaPage)}{frameInfo}");
Console.WriteLine("수집 시작... (Ctrl+C로 중단)");
Console.WriteLine();

using var httpClient = !string.IsNullOrWhiteSpace(apiKey) ? new HttpClient { Timeout = TimeSpan.FromSeconds(15) } : null;

var collected = existingCount;
var visionErrors = 0;
const int maxVisionErrors = 10;
var consecutiveDuplicates = 0;
const int maxConsecutiveDuplicates = 10;
string? lastImageHash = null;

var cts = new CancellationTokenSource();
Console.CancelKeyPress += (_, e) =>
{
    e.Cancel = true;
    cts.Cancel();
    Console.WriteLine("\n중단 요청됨. 마지막 샘플 저장 후 종료합니다.");
};

while (collected < targetCount && !cts.IsCancellationRequested)
{
    try
    {
        var imgCount = await imageLocator.CountAsync();
        if (imgCount == 0)
        {
            Console.WriteLine("  CAPTCHA 이미지가 사라졌습니다. 재탐색 대기 중...");
            await Task.Delay(2000, cts.Token);
            continue;
        }

        var imgBytes = await imageLocator.First.ScreenshotAsync();
        if (imgBytes.Length < 100)
        {
            Console.WriteLine("  이미지가 너무 작습니다. 새로고침 후 재시도...");
            await TryRefreshAsync(captchaPage, captchaFrame, cts.Token);
            await Task.Delay(500, cts.Token);
            continue;
        }

        var currentHash = Convert.ToHexString(MD5.HashData(imgBytes));
        if (currentHash == lastImageHash)
        {
            consecutiveDuplicates++;
            if (consecutiveDuplicates >= maxConsecutiveDuplicates)
            {
                Console.Error.WriteLine($"  ⚠ 연속 {maxConsecutiveDuplicates}회 동일 이미지 감지 — 새로고침 실패 가능성. 1초 대기 후 재시도...");
                await Task.Delay(1000, cts.Token);
                consecutiveDuplicates = 0;
            }
            else
            {
                Console.WriteLine($"  중복 이미지 감지 ({consecutiveDuplicates}/{maxConsecutiveDuplicates}). 새로고침 재시도...");
            }
            await TryRefreshAsync(captchaPage, captchaFrame, cts.Token);
            await Task.Delay(500, cts.Token);
            continue;
        }

        consecutiveDuplicates = 0;
        lastImageHash = currentHash;

        collected++;
        var filename = $"captcha_{collected:D4}.png";
        var filePath = Path.Combine(samplesDir, filename);
        await File.WriteAllBytesAsync(filePath, imgBytes, cts.Token);

        var label = string.Empty;
        if (httpClient is not null && visionErrors < maxVisionErrors)
        {
            label = await CallVisionApiAsync(httpClient, apiKey!, imgBytes);
            if (string.IsNullOrEmpty(label))
                visionErrors++;
        }

        if (!string.IsNullOrEmpty(label))
            existingLabels[filename] = label;

        var labelDisplay = string.IsNullOrEmpty(label) ? "N/A" : label;
        Console.WriteLine($"[{collected}/{targetCount}] {filename}  GT={labelDisplay}  ({imgBytes.Length} bytes)");

        if (collected % 50 == 0)
            await SaveGroundTruthAsync(groundTruthPath, existingLabels);

        await TryRefreshAsync(captchaPage, captchaFrame, cts.Token);
        await Task.Delay(300, cts.Token);
    }
    catch (OperationCanceledException)
    {
        break;
    }
    catch (PlaywrightException ex)
    {
        Console.Error.WriteLine($"  Playwright 오류: {ex.Message}");
        await Task.Delay(2000, cts.Token);
    }
}

await SaveGroundTruthAsync(groundTruthPath, existingLabels);
Console.WriteLine($"\n수집 완료: {collected}개 샘플, {existingLabels.Count}개 레이블");
Console.WriteLine($"저장 위치: {samplesDir}");

return 0;

static async Task TryRefreshAsync(IPage page, IFrame? frame, CancellationToken ct)
{
    var jsTarget = frame as object ?? page;

    try
    {
        bool jsResult;
        if (frame is not null)
        {
            jsResult = await frame.EvaluateAsync<bool>(@"() => {
                if (typeof fnCapchaRefresh === 'function') { fnCapchaRefresh(); return true; }
                if (typeof fnRefresh === 'function') { fnRefresh(); return true; }
                if (typeof captchaRefresh === 'function') { captchaRefresh(); return true; }
                if (typeof refreshCaptcha === 'function') { refreshCaptcha(); return true; }
                return false;
            }");
        }
        else
        {
            jsResult = await page.EvaluateAsync<bool>(@"() => {
                if (typeof fnCapchaRefresh === 'function') { fnCapchaRefresh(); return true; }
                if (typeof fnRefresh === 'function') { fnRefresh(); return true; }
                if (typeof captchaRefresh === 'function') { captchaRefresh(); return true; }
                if (typeof refreshCaptcha === 'function') { refreshCaptcha(); return true; }
                return false;
            }");
        }
        if (jsResult) { await Task.Delay(300, ct); return; }
    }
    catch { }

    try
    {
        var refreshLoc = frame is not null ? frame.Locator(RefreshSelector) : page.Locator(RefreshSelector);
        if (await refreshLoc.CountAsync() > 0)
        {
            await refreshLoc.First.ClickAsync(new LocatorClickOptions { Timeout = 1000, Force = true });
            await Task.Delay(300, ct);
            return;
        }
    }
    catch { }

    try
    {
        var jsCode = @"() => {
            var imgs = document.querySelectorAll('#imgCaptcha, #captchaImg, img[src*=""captcha"" i], img[src*=""cap_img"" i]');
            for (var img of imgs) {
                var src = img.src || '';
                if (src.startsWith('data:')) continue;
                img.src = src.split('?')[0] + '?t=' + Date.now();
            }
            return true;
        }";
        if (frame is not null)
            await frame.EvaluateAsync(jsCode);
        else
            await page.EvaluateAsync(jsCode);
        await Task.Delay(300, ct);
    }
    catch { }
}

static async Task SaveGroundTruthAsync(string path, Dictionary<string, string> labels)
{
    var lines = labels.OrderBy(kv => kv.Key).Select(kv => $"{kv.Key},{kv.Value}");
    await File.WriteAllLinesAsync(path, lines);
}

static string SafeUrl(IPage page)
{
    try { return page.Url.Length > 80 ? page.Url[..80] + "..." : page.Url; }
    catch { return "(unknown)"; }
}

static async Task<string> CallVisionApiAsync(HttpClient httpClient, string apiKey, byte[] imageBytes)
{
    try
    {
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
            return string.Empty;

        var responseJson = await response.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(responseJson);
        var contentArray = doc.RootElement.GetProperty("content");
        foreach (var block in contentArray.EnumerateArray())
        {
            if (block.GetProperty("type").GetString() == "text")
            {
                var raw = block.GetProperty("text").GetString()?.Trim() ?? string.Empty;
                var filtered = new string(raw.Where(char.IsLetterOrDigit).Select(char.ToUpperInvariant).ToArray());
                if (Regex.IsMatch(filtered, "^[A-Z0-9]{6}$"))
                    return filtered;
            }
        }
    }
    catch { }
    return string.Empty;
}
