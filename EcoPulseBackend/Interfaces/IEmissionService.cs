namespace EcoPulseBackend.Interfaces;

public interface IEmissionService
{
    public IGasolineGeneratorService GasolineGeneratorService { get; }

    public IReservoirsService ReservoirsService { get; }

    public IDuringMetalMachiningService DuringMetalMachiningService { get; }

    public IDuringWeldingOperationsService DuringWeldingOperationsService { get; }

    public IMaximumSingleService MaximumSingleService { get; }

    public IVehicleFlowService VehicleFlowService { get; }

    public ITrafficLightQueueService TrafficLightQueueService { get; }

    public IOpenCoalWarehouseService OpenCoalWarehouseService { get; }
}