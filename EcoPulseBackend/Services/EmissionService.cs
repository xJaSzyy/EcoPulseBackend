using EcoPulseBackend.Enums;
using EcoPulseBackend.Extensions;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DangerZone;
using EcoPulseBackend.Models.DuringMetalMachining;
using EcoPulseBackend.Models.DuringWeldingOperations;
using EcoPulseBackend.Models.GasolineGenerator;
using EcoPulseBackend.Models.MaximumSingle;
using EcoPulseBackend.Models.OpenCoalWarehouse;
using EcoPulseBackend.Models.Reservoirs;
using EcoPulseBackend.Models.TrafficLightQueue;
using EcoPulseBackend.Models.TrafficLightQueueEmissionSource;
using EcoPulseBackend.Models.VehicleFlow;
using EcoPulseBackend.Models.VehicleFlowEmissionSource;

namespace EcoPulseBackend.Services;

public class EmissionService : IEmissionService
{
    public IGasolineGeneratorService GasolineGeneratorService { get; }
    public IReservoirsService ReservoirsService { get; }
    public IDuringMetalMachiningService DuringMetalMachiningService { get; }
    public IDuringWeldingOperationsService DuringWeldingOperationsService { get; }
    public IMaximumSingleService MaximumSingleService { get; }
    public IVehicleFlowService VehicleFlowService { get; }
    public ITrafficLightQueueService TrafficLightQueueService { get; }
    public IOpenCoalWarehouseService OpenCoalWarehouseService { get; }

    public EmissionService(IGasolineGeneratorService gasolineGeneratorService, IReservoirsService reservoirsService,
        IDuringMetalMachiningService duringMetalMachiningService, IDuringWeldingOperationsService duringWeldingOperationsService,
        IMaximumSingleService maximumSingleService, IVehicleFlowService vehicleFlowService,
        ITrafficLightQueueService trafficLightQueueService, IOpenCoalWarehouseService openCoalWarehouseService)
    {
        GasolineGeneratorService = gasolineGeneratorService;
        ReservoirsService = reservoirsService;
        DuringMetalMachiningService = duringMetalMachiningService;
        DuringWeldingOperationsService = duringWeldingOperationsService;
        MaximumSingleService = maximumSingleService;
        VehicleFlowService = vehicleFlowService;
        TrafficLightQueueService = trafficLightQueueService;
        OpenCoalWarehouseService = openCoalWarehouseService;
    }
}