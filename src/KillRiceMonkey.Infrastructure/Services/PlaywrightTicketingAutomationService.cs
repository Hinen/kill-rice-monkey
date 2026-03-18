using Microsoft.Extensions.Logging;
using Microsoft.Playwright;
using OpenCvSharp;
using Polly;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;
using KillRiceMonkey.Application.Abstractions;
using KillRiceMonkey.Application.Models;
using StepGroup = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.StepGroup;
using StepTemplate = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.StepTemplate;
using TemplateMetadata = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.TemplateMetadata;
using MatchHit = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.MatchHit;
using PriorityCandidate = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.PriorityCandidate;
using TemplateBounds = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.TemplateBounds;

namespace KillRiceMonkey.Infrastructure.Services;

public sealed partial class PlaywrightTicketingAutomationService : ITicketingAutomationService, IAsyncDisposable
{
    private static readonly Regex NolRoundPattern = new(@"^\D*(?<round>\d{1,2})\s*(?:회차|회|희|히|외)?\s*(?<time>\d{1,2}(?::|\.|,)?\d{2})", RegexOptions.Compiled);
    private const int SeatSelectionOffset = 1;
    private const int Yes24LegendPaddingX = 8;
    private const double Yes24LegendSearchStartRatio = 0.70;
    private const double Yes24LegendMinIgnoreRatio = 0.55;
    private const double Yes24LegendThresholdDelta = 0.08;
    private const double Yes24LegendFallbackIgnoreRatio = 0.72;
    private const double Yes24LegendMaxIgnoreRatio = 0.92;
    private const int Yes24SeatColorSampleRadius = 6;
    private const double Yes24SeatMinSaturation = 35;

    private const string NolRemoteDebugLaunchUrl = "https://tickets.interpark.com/";
    private const string NolCdpEndpoint = "http://127.0.0.1:9222/";
    private const string MelonRemoteDebugLaunchUrl = "https://ticket.melon.com/";
    private const int MelonRemoteDebugPort = 9223;
    private const string MelonCdpEndpoint = "http://127.0.0.1:9223/";
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
    private readonly ILogger<PlaywrightTicketingAutomationService> _logger;
    private readonly PlaywrightRuntime _runtime;
    private readonly ResiliencePipeline _pipeline;
    private readonly SemaphoreSlim _nolBrowserLock = new(1, 1);
    private readonly SemaphoreSlim _melonBrowserLock = new(1, 1);
    private IBrowser? _preparedNolConnectedBrowser;
    private IPage? _preparedNolPage;
    private IBrowser? _preparedMelonConnectedBrowser;
    private IPage? _preparedMelonPage;
    private bool _popupClosedDuringPrepare;
    private bool _melonPopupClosedDuringPrepare;

    public PlaywrightTicketingAutomationService(
        ILogger<PlaywrightTicketingAutomationService> logger,
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
        _nolBrowserLock.Dispose();
        _melonBrowserLock.Dispose();
    }

    public async Task<bool> IsNolRemoteDebugBrowserAvailableAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return await PlaywrightRuntime.IsCdpEndpointAvailableAsync(NolCdpEndpoint, cancellationToken);
    }

    public async Task<bool> IsNolAutomationPreparedAsync(CancellationToken cancellationToken)
    {
        return await TryGetPreparedNolPageAsync(cancellationToken) is not null;
    }

    public async Task<bool> IsMelonRemoteDebugBrowserAvailableAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return await PlaywrightRuntime.IsCdpEndpointAvailableAsync(MelonCdpEndpoint, cancellationToken);
    }

    public async Task<bool> IsMelonAutomationPreparedAsync(CancellationToken cancellationToken)
    {
        return await TryGetPreparedMelonPageAsync(cancellationToken) is not null;
    }

    public async Task<string> LaunchNolRemoteDebugBrowserAsync(CancellationToken cancellationToken)
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

    public async Task<string> LaunchMelonRemoteDebugBrowserAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (await PlaywrightRuntime.IsCdpEndpointAvailableAsync(MelonCdpEndpoint, cancellationToken))
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
            if (await PlaywrightRuntime.IsCdpEndpointAvailableAsync(MelonCdpEndpoint, cancellationToken))
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
        return $"NOL 준비 완료: {PlaywrightRuntime.SafePageUrl(page)}";
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
        return $"Melon 준비 완료: {PlaywrightRuntime.SafePageUrl(page)}";
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

            using var frame = PlaywrightRuntime.CaptureScreen();
            using var grayFrame = PlaywrightRuntime.ToGray(frame.Image);
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
                clickMatch = PlaywrightRuntime.TryFindMatch(grayFrame, clickableTemplates, threshold, out clickBestScore);
            }

            bestScore = Math.Max(bestScore, clickBestScore);
            if (clickMatch is not null)
            {
                PlaywrightRuntime.ClickAt(frame.OffsetX + clickMatch.Value.X, frame.OffsetY + clickMatch.Value.Y);

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

            await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
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
            var match = PlaywrightRuntime.StepPattern.Match(fileName);
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

            templates.Add(new StepTemplate(metadata.State, metadata.Priority, metadata.IsViewMask, PlaywrightRuntime.BuildScaledTemplates(image)));
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
        var prefix = $"{PlaywrightRuntime.EmbeddedTemplatePrefix}{templateType}.";

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
            var match = PlaywrightRuntime.StepPattern.Match(fileName);
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

            templates.Add(new StepTemplate(metadata.State, metadata.Priority, metadata.IsViewMask, PlaywrightRuntime.BuildScaledTemplates(encoded)));
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

            using var frame = PlaywrightRuntime.CaptureScreen();
            using var grayFrame = PlaywrightRuntime.ToGray(frame.Image);

            if (activeTemplates.Count > 0 && PlaywrightRuntime.TryFindMatch(grayFrame, activeTemplates, threshold, out _) is not null)
            {
                return true;
            }

            if (normalTemplates.Count > 0 && PlaywrightRuntime.TryFindMatch(grayFrame, normalTemplates, threshold, out _) is null)
            {
                return true;
            }

            await Task.Delay(PlaywrightRuntime.PollDelayMilliseconds, cancellationToken);
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


}
