using EcoPulseBackend.Contexts;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models.DuringMetalMachining;
using EcoPulseBackend.Models.DuringWeldingOperations;
using EcoPulseBackend.Models.GasolineGenerator;
using EcoPulseBackend.Models.MaximumSingle;
using EcoPulseBackend.Models.OpenCoalWarehouse;
using EcoPulseBackend.Models.Reservoirs;
using EcoPulseBackend.Models.TrafficLightQueue;
using EcoPulseBackend.Models.VehicleFlow;
using Microsoft.AspNetCore.Mvc;

namespace EcoPulseBackend.Controllers;

[ApiController]
public class EmissionController : ControllerBase
{
    private readonly IEmissionService _emissionService;
    private readonly ApplicationDbContext _dbContext;
    private readonly ILogger<EmissionController> _logger;

    public EmissionController(IEmissionService emissionService, ApplicationDbContext dbContext, ILogger<EmissionController> logger)
    {
        _emissionService = emissionService;
        _dbContext = dbContext;
        _logger = logger;
    }

    [HttpPost("emission/gasoline-generator")]
    public IActionResult CalculateGasolineGenerator([FromBody] GasolineGeneratorEmissionsCalculateModel model)
    {
        var result = _emissionService.GasolineGeneratorService.CalculateGasolineGeneratorEmissionsBatch(model);

        return Ok(result);
    }

    [HttpPost("emission/reservoirs")]
    public IActionResult CalculateReservoirs([FromBody] ReservoirsEmissionsCalculateModel model)
    {
        var result = _emissionService.ReservoirsService.CalculateReservoirsEmissionsBatch(model);

        return Ok(result.Emissions);
    }
    
    [HttpPost("emission/during-metal-machining")]
    public IActionResult CalculateDuringMetalMachining([FromBody] DuringMetalMachiningEmissionsCalculateModel model)
    { 
        var result = _emissionService.DuringMetalMachiningService.CalculateDuringMetalMachiningEmissionsBatch(model);

        return Ok(result);
    }
    
    [HttpPost("emission/during-welding-operations")]
    public IActionResult CalculateDuringWeldingOperations([FromBody] DuringWeldingOperationsEmissionsCalculateModel model)
    {
        var result = _emissionService.DuringWeldingOperationsService.CalculateDuringWeldingOperationsEmissionsBatch(model);

        return Ok(result.Emissions);
    }
    
    [HttpPost("emission/maximum-single")]
    public IActionResult CalculateMaximumSingleEmissions([FromBody] MaximumSingleEmissionsCalculateModel model)
    {
        var result = _emissionService.MaximumSingleService.CalculateMaximumSingleEmissions(model);

        return Ok(result);
    }
    
    [HttpPost("emission/vehicle-flow")]
    public IActionResult CalculateVehicleFlowEmissions([FromBody] VehicleFlowEmissionsCalculateModel model)
    {
        var result = _emissionService.VehicleFlowService.CalculateVehicleFlowEmissionsBatch(model);

        return Ok(result);
    }
    
    [HttpPost("emission/traffic-light-queue")]
    public IActionResult CalculateTrafficLightQueueEmissions([FromBody] TrafficLightQueueEmissionsCalculateModel model)
    {
        var result = _emissionService.TrafficLightQueueService.CalculateTrafficLightQueueEmissionsBatch(model);

        return Ok(result);
    }

    [HttpPost("emission/open-coal-warehouse")]
    public IActionResult CalculateOpenCoalWarehouseEmissions([FromBody] OpenCoalWarehouseEmissionsCalculateModel model)
    {
        var result = _emissionService.OpenCoalWarehouseService.CalculateOpenCoalWarehouseEmissions(model);

        return Ok(result);
    }
}