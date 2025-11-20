using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Services;

namespace EcoPulseBackend;

public static class DependencyInjection
{
    public static void AddServices(this IServiceCollection services)
    {
        services.AddScoped<IEmissionService, EmissionService>();
        services.AddScoped<IExportService, ExportService>();
    }
}