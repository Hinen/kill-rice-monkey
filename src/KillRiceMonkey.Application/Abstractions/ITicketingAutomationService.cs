using KillRiceMonkey.Application.Models;

namespace KillRiceMonkey.Application.Abstractions;

public interface ITicketingAutomationService
{
    Task<AutomationRunResult> RunAsync(TicketingJobRequest request, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken);
    Task<AutomationRunResult> RunAsync(TicketingJobRequest request, CancellationToken cancellationToken);
    Task<bool> IsNolRemoteDebugBrowserAvailableAsync(CancellationToken cancellationToken);
    Task<bool> IsNolAutomationPreparedAsync(CancellationToken cancellationToken);
    Task<string> LaunchNolRemoteDebugBrowserAsync(CancellationToken cancellationToken);
    Task<string> PrepareNolAutomationAsync(CancellationToken cancellationToken);
    Task<bool> IsNolPageReadyAsync(CancellationToken cancellationToken);
    Task<bool> IsMelonRemoteDebugBrowserAvailableAsync(CancellationToken cancellationToken);
    Task<bool> IsMelonAutomationPreparedAsync(CancellationToken cancellationToken);
    Task<string> LaunchMelonRemoteDebugBrowserAsync(CancellationToken cancellationToken);
    Task<string> PrepareMelonAutomationAsync(CancellationToken cancellationToken);
    Task<bool> IsMelonPageReadyAsync(CancellationToken cancellationToken);
}
