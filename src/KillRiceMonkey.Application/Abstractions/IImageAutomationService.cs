using KillRiceMonkey.Application.Models;

namespace KillRiceMonkey.Application.Abstractions;

public interface IImageAutomationService
{
    Task<AutomationRunResult> RunAsync(TicketingJobRequest request, CancellationToken cancellationToken);
    Task<AutomationRunResult> RunAsync(TicketingJobRequest request, IProgress<AutomationProgress>? progress, CancellationToken cancellationToken);
}
