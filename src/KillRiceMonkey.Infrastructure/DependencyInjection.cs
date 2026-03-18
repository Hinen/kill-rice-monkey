using Microsoft.Extensions.DependencyInjection;
using KillRiceMonkey.Application.Abstractions;
using KillRiceMonkey.Infrastructure.Services;

namespace KillRiceMonkey.Infrastructure;

public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructure(this IServiceCollection services)
    {
        services.AddSingleton<PlaywrightRuntime>();
        services.AddSingleton<IImageAutomationService, ImageAutomationService>();
        services.AddSingleton<INolAutomationService, NolAutomationService>();
        services.AddSingleton<IMelonAutomationService, MelonAutomationService>();
        services.AddSingleton<ITicketingAutomationService, PlaywrightTicketingAutomationService>();
        return services;
    }
}
