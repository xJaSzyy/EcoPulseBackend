using EcoPulseBackend.Contexts;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Services;
using Microsoft.EntityFrameworkCore;

namespace EcoPulseBackend;

public static class DependencyInjection
{
    public static void AddServices(this IServiceCollection services)
    {
        services.AddScoped<IEmissionService, EmissionService>();
        services.AddScoped<IExportService, ExportService>();
    }

    public static void AddDatabase(this IServiceCollection services)
    {
        services.AddDbContext<ApplicationDbContext>(options =>
            options.UseNpgsql(Environment.GetEnvironmentVariable("ConnectionStrings__DefaultConnection")));
    }
}