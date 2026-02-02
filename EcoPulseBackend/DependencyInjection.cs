using EcoPulseBackend.Contexts;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Services;
using Microsoft.EntityFrameworkCore;

namespace EcoPulseBackend;

public static class DependencyInjection
{
    public static void AddServices(this IServiceCollection services)
    {
        services.AddScoped<IGasolineGeneratorService, GasolineGeneratorService>();
        services.AddScoped<IReservoirsService, ReservoirsService>();
        services.AddScoped<IDuringMetalMachiningService, DuringMetalMachiningService>();
        services.AddScoped<IDuringWeldingOperationsService, DuringWeldingOperationsService>();
        services.AddScoped<IMaximumSingleService, MaximumSingleService>();
        services.AddScoped<IVehicleFlowService, VehicleFlowService>();
        services.AddScoped<ITrafficLightQueueService, TrafficLightQueueService>();
        services.AddScoped<IOpenCoalWarehouseService, OpenCoalWarehouseService>();
        services.AddScoped<IEmissionService, EmissionService>();
        services.AddScoped<IExportService, ExportService>();
        services.AddScoped<IRecommendationService, RecommendationService>();
    }

    public static void AddDatabase(this IServiceCollection services, ConfigurationManager configuration)
    {
        var connectionString = configuration.GetConnectionString("DefaultConnection");

        services.AddDbContext<ApplicationDbContext>(options =>
            options.UseNpgsql(connectionString));
    }
}