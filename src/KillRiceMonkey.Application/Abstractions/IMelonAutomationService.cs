using KillRiceMonkey.Application.Models;

namespace KillRiceMonkey.Application.Abstractions;

public interface IMelonAutomationService
{
    Task<AutomationRunResult> RunAsync(TicketingJobRequest request, CancellationToken cancellationToken);
    Task<AutomationRunResult> RunAsync(TicketingJobRequest request, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken);
    Task<bool> IsRemoteDebugBrowserAvailableAsync(CancellationToken cancellationToken);
    Task<bool> IsAutomationPreparedAsync(CancellationToken cancellationToken);
    Task<string> LaunchRemoteDebugBrowserAsync(CancellationToken cancellationToken);
    Task<string> PrepareAutomationAsync(CancellationToken cancellationToken);
    Task<bool> IsPageReadyAsync(CancellationToken cancellationToken);
}
