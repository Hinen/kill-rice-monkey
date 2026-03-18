using KillRiceMonkey.Application.Abstractions;
using KillRiceMonkey.Application.Models;

namespace KillRiceMonkey.Infrastructure.Services;

public sealed class PlaywrightTicketingAutomationService : ITicketingAutomationService, IAsyncDisposable
{
    private readonly IImageAutomationService _imageAutomationService;
    private readonly INolAutomationService _nolAutomationService;
    private readonly IMelonAutomationService _melonAutomationService;

    public PlaywrightTicketingAutomationService(
        IImageAutomationService imageAutomationService,
        INolAutomationService nolAutomationService,
        IMelonAutomationService melonAutomationService)
    {
        _imageAutomationService = imageAutomationService;
        _nolAutomationService = nolAutomationService;
        _melonAutomationService = melonAutomationService;
    }

    public Task<AutomationRunResult> RunAsync(TicketingJobRequest request, CancellationToken cancellationToken)
        => RunAsync(request, null, cancellationToken);

    public Task<AutomationRunResult> RunAsync(TicketingJobRequest request, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken)
    {
        return request.TemplateType switch
        {
            TicketingTemplateType.Nol => _nolAutomationService.RunAsync(request, progress, cancellationToken),
            TicketingTemplateType.Melon => _melonAutomationService.RunAsync(request, progress, cancellationToken),
            _ => _imageAutomationService.RunAsync(request, progress, cancellationToken)
        };
    }

    public ValueTask DisposeAsync() => ValueTask.CompletedTask;

    public Task<bool> IsNolRemoteDebugBrowserAvailableAsync(CancellationToken cancellationToken)
        => _nolAutomationService.IsRemoteDebugBrowserAvailableAsync(cancellationToken);

    public Task<bool> IsNolAutomationPreparedAsync(CancellationToken cancellationToken)
        => _nolAutomationService.IsAutomationPreparedAsync(cancellationToken);

    public Task<bool> IsMelonRemoteDebugBrowserAvailableAsync(CancellationToken cancellationToken)
        => _melonAutomationService.IsRemoteDebugBrowserAvailableAsync(cancellationToken);

    public Task<bool> IsMelonAutomationPreparedAsync(CancellationToken cancellationToken)
        => _melonAutomationService.IsAutomationPreparedAsync(cancellationToken);

    public Task<string> LaunchNolRemoteDebugBrowserAsync(CancellationToken cancellationToken)
        => _nolAutomationService.LaunchRemoteDebugBrowserAsync(cancellationToken);

    public Task<string> LaunchMelonRemoteDebugBrowserAsync(CancellationToken cancellationToken)
        => _melonAutomationService.LaunchRemoteDebugBrowserAsync(cancellationToken);

    public Task<string> PrepareNolAutomationAsync(CancellationToken cancellationToken)
        => _nolAutomationService.PrepareAutomationAsync(cancellationToken);

    public Task<string> PrepareMelonAutomationAsync(CancellationToken cancellationToken)
        => _melonAutomationService.PrepareAutomationAsync(cancellationToken);

    public Task<bool> IsNolPageReadyAsync(CancellationToken cancellationToken)
        => _nolAutomationService.IsPageReadyAsync(cancellationToken);

    public Task<bool> IsMelonPageReadyAsync(CancellationToken cancellationToken)
        => _melonAutomationService.IsPageReadyAsync(cancellationToken);
}
