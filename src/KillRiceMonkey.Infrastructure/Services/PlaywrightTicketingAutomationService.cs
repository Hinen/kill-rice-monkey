using Microsoft.Extensions.Logging;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Polly;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using KillRiceMonkey.Application.Abstractions;
using KillRiceMonkey.Application.Models;

namespace KillRiceMonkey.Infrastructure.Services;

public sealed class PlaywrightTicketingAutomationService : ITicketingAutomationService
{
    private static readonly Regex StepPattern = new("^(?<step>\\d+)-(?<state>normal|active)\\.png$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private const int SmXvirtualscreen = 76;
    private const int SmYvirtualscreen = 77;
    private const int SmCxvirtualscreen = 78;
    private const int SmCyvirtualscreen = 79;
    private const int InputMouse = 0;
    private const uint MouseeventfLeftdown = 0x0002;
    private const uint MouseeventfLeftup = 0x0004;

    private readonly ILogger<PlaywrightTicketingAutomationService> _logger;
    private readonly ResiliencePipeline _pipeline;

    public PlaywrightTicketingAutomationService(ILogger<PlaywrightTicketingAutomationService> logger)
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

    public async Task<AutomationRunResult> RunAsync(TicketingJobRequest request, CancellationToken cancellationToken)
    {
        return await _pipeline.ExecuteAsync(async token =>
        {
            var resolvedDirectory = Path.GetFullPath(request.ImageDirectory, AppContext.BaseDirectory);
            if (!Directory.Exists(resolvedDirectory))
            {
                return new AutomationRunResult(false, $"이미지 폴더를 찾을 수 없습니다: {resolvedDirectory}", DateTimeOffset.Now);
            }

            var stepGroups = LoadStepGroups(resolvedDirectory);
            try
            {
                if (stepGroups.Count == 0)
                {
                    return new AutomationRunResult(false, "숫자-상태.png 패턴의 이미지가 없습니다.", DateTimeOffset.Now);
                }

                _logger.LogInformation("Automation started. directory={Directory}, stepCount={Count}", resolvedDirectory, stepGroups.Count);

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

    private async Task<(bool IsSuccess, string Message)> WaitAndClickStepAsync(
        StepGroup stepGroup,
        double threshold,
        int timeoutSeconds,
        CancellationToken cancellationToken)
    {
        var started = DateTimeOffset.UtcNow;
        var activeTemplates = stepGroup.Templates.Where(x => x.State == "active").ToList();
        var normalTemplates = stepGroup.Templates.Where(x => x.State == "normal").ToList();

        while (DateTimeOffset.UtcNow - started < TimeSpan.FromSeconds(timeoutSeconds))
        {
            cancellationToken.ThrowIfCancellationRequested();

            using var frame = CaptureScreen();
            var activeMatch = TryFindMatch(frame.Image, activeTemplates, threshold);
            if (activeMatch is not null)
            {
                return (true, $"{stepGroup.Step}단계 이미 active 상태 감지 (score={activeMatch.Value.Score:F3})");
            }

            var normalMatch = TryFindMatch(frame.Image, normalTemplates, threshold);
            if (normalMatch is not null)
            {
                ClickAt(frame.OffsetX + normalMatch.Value.X, frame.OffsetY + normalMatch.Value.Y);
                var transitioned = await WaitForTransitionAsync(
                    activeTemplates,
                    normalTemplates,
                    threshold,
                    timeoutSeconds,
                    started,
                    cancellationToken);

                if (transitioned)
                {
                    return (true, $"{stepGroup.Step}단계 클릭 성공 ({normalMatch.Value.State}, score={normalMatch.Value.Score:F3})");
                }

                return (false, $"{stepGroup.Step}단계 클릭 후 상태 전환 확인 실패");
            }

            await Task.Delay(120, cancellationToken);
        }

        return (false, $"{stepGroup.Step}단계 버튼 탐지 실패 (timeout {timeoutSeconds}s)");
    }

    private static IReadOnlyList<StepGroup> LoadStepGroups(string directory)
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

            var state = match.Groups["state"].Value.ToLowerInvariant();
            var image = Cv2.ImRead(filePath, ImreadModes.Color);
            if (image.Empty())
            {
                continue;
            }

            if (!map.TryGetValue(step, out var templates))
            {
                templates = new List<StepTemplate>();
                map[step] = templates;
            }

            if (templates.Any(x => x.State == state))
            {
                image.Dispose();
                throw new InvalidOperationException($"중복 상태 이미지가 존재합니다: {step}-{state}");
            }

            templates.Add(new StepTemplate(state, image));
        }

        return map
            .OrderBy(kvp => kvp.Key)
            .Select(kvp => new StepGroup(kvp.Key, kvp.Value))
            .ToList();
    }

    private static MatchHit? TryFindMatch(Mat screenshot, IReadOnlyList<StepTemplate> templates, double threshold)
    {
        foreach (var template in templates)
        {
            if (screenshot.Width < template.Image.Width || screenshot.Height < template.Image.Height)
            {
                continue;
            }

            using var result = new Mat();
            Cv2.MatchTemplate(screenshot, template.Image, result, TemplateMatchModes.CCoeffNormed);
            Cv2.MinMaxLoc(result, out _, out var maxValue, out _, out var maxLocation);
            if (maxValue >= threshold)
            {
                return new MatchHit(
                    template.State,
                    maxLocation.X + (template.Image.Width / 2),
                    maxLocation.Y + (template.Image.Height / 2),
                    maxValue);
            }
        }

        return null;
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
            if (activeTemplates.Count > 0 && TryFindMatch(frame.Image, activeTemplates, threshold) is not null)
            {
                return true;
            }

            if (normalTemplates.Count > 0 && TryFindMatch(frame.Image, normalTemplates, threshold) is null)
            {
                return true;
            }

            await Task.Delay(120, cancellationToken);
        }

        return false;
    }

    private static void DisposeStepGroups(IEnumerable<StepGroup> groups)
    {
        foreach (var group in groups)
        {
            foreach (var template in group.Templates)
            {
                template.Image.Dispose();
            }
        }
    }

    private static ScreenFrame CaptureScreen()
    {
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

    private readonly record struct MatchHit(string State, int X, int Y, double Score);

    private sealed record StepTemplate(string State, Mat Image);
    private sealed record StepGroup(int Step, IReadOnlyList<StepTemplate> Templates);
}
