using KillRiceMonkey.Application.Models;

namespace KillRiceMonkey.Application.Abstractions;

public interface ITicketingAutomationService
{
    Task<AutomationRunResult> RunAsync(TicketingJobRequest request, CancellationToken cancellationToken);
}
