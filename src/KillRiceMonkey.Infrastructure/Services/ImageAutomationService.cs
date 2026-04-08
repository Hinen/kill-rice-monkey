using Microsoft.Extensions.Logging;
using OpenCvSharp;
using Polly;
using System.Text.RegularExpressions;
using KillRiceMonkey.Application.Abstractions;
using KillRiceMonkey.Application.Models;
using StepGroup = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.StepGroup;
using StepTemplate = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.StepTemplate;
using TemplateMetadata = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.TemplateMetadata;
using MatchHit = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.MatchHit;
using PriorityCandidate = KillRiceMonkey.Infrastructure.Services.PlaywrightRuntime.PriorityCandidate;

namespace KillRiceMonkey.Infrastructure.Services;

public sealed class ImageAutomationService : IImageAutomationService
{
    private const int SeatSelectionOffset = 1;

    private readonly ILogger<ImageAutomationService> _logger;
    private readonly ResiliencePipeline _pipeline;

    public ImageAutomationService(ILogger<ImageAutomationService> logger)
    {
        _logger = logger;
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

                        var found = await WaitAndClickStepAsync(stepGroup, request.MatchThreshold, request.StepTimeoutSeconds, token);
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

    private static async Task<(bool IsSuccess, string Message)> WaitAndClickStepAsync(
        StepGroup stepGroup,
        double threshold,
        int timeoutSeconds,
        CancellationToken cancellationToken)
    {
        var started = DateTimeOffset.UtcNow;
        var activeTemplates = stepGroup.Templates.Where(x => x.State == "active").ToList();
        var normalTemplates = stepGroup.Templates.Where(x => x.State == "normal" && !x.IsViewMask).ToList();
        var singleTemplates = stepGroup.Templates.Where(x => x.State == "single" && !x.IsViewMask).ToList();
        var hasPriorityTemplates = singleTemplates.Any(x => x.Priority is not null);

        if (normalTemplates.Count == 0 && singleTemplates.Count == 0)
        {
            return (false, $"{stepGroup.Step}단계 클릭 가능한 이미지가 없습니다.");
        }

        var bestScore = double.NegativeInfinity;

        while (DateTimeOffset.UtcNow - started < TimeSpan.FromSeconds(timeoutSeconds))
        {
            cancellationToken.ThrowIfCancellationRequested();

            using var frame = PlaywrightRuntime.CaptureScreen();
            using var grayFrame = PlaywrightRuntime.ToGray(frame.Image);

            var clickableTemplates = normalTemplates.Count > 0 ? normalTemplates : singleTemplates;
            MatchHit? clickMatch;
            double clickBestScore;

            if (hasPriorityTemplates && normalTemplates.Count == 0)
            {
                clickMatch = TryFindPrioritySeatMatch(grayFrame, clickableTemplates, threshold, out clickBestScore);
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
        var assembly = typeof(ImageAutomationService).Assembly;
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

    private static MatchHit? TryFindPrioritySeatMatch(
        Mat grayScreenshot,
        IReadOnlyList<StepTemplate> templates,
        double threshold,
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

            CollectPriorityCandidates(grayScreenshot, priorityGroup, threshold, candidates, fastPathOnly: true, ref bestScore);
            if (candidates.Count == 0)
            {
                CollectPriorityCandidates(grayScreenshot, priorityGroup, threshold, candidates, fastPathOnly: false, ref bestScore);
            }

            if (candidates.Count == 0)
            {
                continue;
            }

            var ordered = DeduplicateCandidates(candidates)
                .OrderBy(x => x.Y)
                .ThenBy(x => x.X)
                .ToList();

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

                CollectMatchCandidates(result, scaledTemplate.Width, scaledTemplate.Height, threshold, candidates);
            }
        }
    }

    private static void CollectMatchCandidates(
        Mat result,
        int templateWidth,
        int templateHeight,
        double threshold,
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
