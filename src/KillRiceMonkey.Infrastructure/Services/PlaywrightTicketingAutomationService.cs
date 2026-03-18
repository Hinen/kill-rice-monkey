using Microsoft.Extensions.Logging;
using Microsoft.Playwright;
using Polly;
using System.Diagnostics;
using KillRiceMonkey.Application.Abstractions;
using KillRiceMonkey.Application.Models;

namespace KillRiceMonkey.Infrastructure.Services;

public sealed partial class PlaywrightTicketingAutomationService : ITicketingAutomationService, IAsyncDisposable
{
    private const int SeatSelectionOffset = 1;
    private const string MelonRemoteDebugLaunchUrl = "https://ticket.melon.com/";
    private const int MelonRemoteDebugPort = 9223;
    private const string MelonCdpEndpoint = "http://127.0.0.1:9223/";
    private static readonly string[] NolBrowserExecutableCandidates =
    [
        @"C:\Program Files\Google\Chrome\Application\chrome.exe",
        @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        @"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    ];
    private readonly ILogger<PlaywrightTicketingAutomationService> _logger;
    private readonly PlaywrightRuntime _runtime;
    private readonly IImageAutomationService _imageAutomationService;
    private readonly INolAutomationService _nolAutomationService;
    private readonly ResiliencePipeline _pipeline;
    private readonly SemaphoreSlim _melonBrowserLock = new(1, 1);
    private IBrowser? _preparedMelonConnectedBrowser;
    private IPage? _preparedMelonPage;
    private bool _melonPopupClosedDuringPrepare;

    public PlaywrightTicketingAutomationService(
        ILogger<PlaywrightTicketingAutomationService> logger,
        PlaywrightRuntime runtime,
        IImageAutomationService imageAutomationService,
        INolAutomationService nolAutomationService)
    {
        _logger = logger;
        _runtime = runtime;
        _imageAutomationService = imageAutomationService;
        _nolAutomationService = nolAutomationService;
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
        if (request.TemplateType == TicketingTemplateType.Nol)
        {
            return await _nolAutomationService.RunAsync(request, progress, cancellationToken);
        }

        if (request.TemplateType != TicketingTemplateType.Melon)
        {
            return await _imageAutomationService.RunAsync(request, progress, cancellationToken);
        }

        try
        {
            return await _pipeline.ExecuteAsync(async token =>
            {
                return await RunMelonAutomationAsync(request, progress, token);
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
        await ReleasePreparedMelonConnectionAsync();
        _melonBrowserLock.Dispose();
    }

    public Task<bool> IsNolRemoteDebugBrowserAvailableAsync(CancellationToken cancellationToken)
        => _nolAutomationService.IsRemoteDebugBrowserAvailableAsync(cancellationToken);

    public Task<bool> IsNolAutomationPreparedAsync(CancellationToken cancellationToken)
        => _nolAutomationService.IsAutomationPreparedAsync(cancellationToken);

    public async Task<bool> IsMelonRemoteDebugBrowserAvailableAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return await PlaywrightRuntime.IsCdpEndpointAvailableAsync(MelonCdpEndpoint, cancellationToken);
    }

    public async Task<bool> IsMelonAutomationPreparedAsync(CancellationToken cancellationToken)
    {
        return await TryGetPreparedMelonPageAsync(cancellationToken) is not null;
    }

    public Task<string> LaunchNolRemoteDebugBrowserAsync(CancellationToken cancellationToken)
        => _nolAutomationService.LaunchRemoteDebugBrowserAsync(cancellationToken);

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

    public Task<string> PrepareNolAutomationAsync(CancellationToken cancellationToken)
        => _nolAutomationService.PrepareAutomationAsync(cancellationToken);

    public Task<bool> IsNolPageReadyAsync(CancellationToken cancellationToken)
        => _nolAutomationService.IsPageReadyAsync(cancellationToken);

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
}
