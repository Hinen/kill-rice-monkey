using Microsoft.Extensions.DependencyInjection;
using TicketingAutoPurchase.Application.Abstractions;
using TicketingAutoPurchase.Infrastructure.Services;

namespace TicketingAutoPurchase.Infrastructure;

public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructure(this IServiceCollection services)
    {
        services.AddHttpClient();
        services.AddSingleton<ITicketingAutomationService, PlaywrightTicketingAutomationService>();
        return services;
    }
}
