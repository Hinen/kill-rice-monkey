using TicketingAutoPurchase.Application.Models;

namespace TicketingAutoPurchase.Application.Abstractions;

public interface ITicketingAutomationService
{
    Task<AutomationRunResult> RunAsync(TicketingJobRequest request, CancellationToken cancellationToken);
}
