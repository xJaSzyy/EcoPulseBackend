using EcoPulseBackend.Enums;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.Calculate;
using Microsoft.AspNetCore.Mvc;

namespace EcoPulseBackend.Controllers;

[ApiController]
public class EmissionController : ControllerBase
{
    private readonly IEmissionService _emissionService;
    private readonly IExportService _exportService;
    private readonly ILogger<EmissionController> _logger;

    public EmissionController(ILogger<EmissionController> logger, IEmissionService emissionService,
        IExportService exportService)
    {
        _emissionService = emissionService;
        _exportService = exportService;
        _logger = logger;
    }
    
    [HttpPost("calculate/gasoline-generator")]
    public IActionResult CalculateGasolineGenerator([FromBody] GasolineGeneratorEmissionsCalculateModel model)
    {
        var result = _emissionService.CalculateGasolineGeneratorEmissionsBatch(model);

        return Ok(result);
    }
    
    [HttpPost("calculate/reservoirs")]
    public IActionResult CalculateReservoirs([FromBody] ReservoirsEmissionsCalculateModel model)
    {
        var result = _emissionService.CalculateReservoirsEmissionsBatch(model);

        return Ok(result.Emissions);
    }
    
    [HttpPost("calculate/during-metal-machining")]
    public IActionResult CalculateDuringMetalMachining([FromBody] DuringMetalMachiningEmissionsCalculateModel model)
    { 
        var result = _emissionService.CalculateDuringMetalMachiningEmissionsBatch(model);

        return Ok(result);
    }
    
    [HttpPost("calculate/during-welding-operations")]
    public IActionResult CalculateDuringWeldingOperations([FromBody] DuringWeldingOperationsEmissionsCalculateModel model)
    {
        var result = _emissionService.CalculateDuringWeldingOperationsEmissionsBatch(model);

        return Ok(result.Emissions);
    }
    
    [HttpPost("calculate/maximum-single")]
    public IActionResult CalculateMaximumSingle([FromBody] MaximumSingleEmissionsCalculateModel model)
    {
        var result = _emissionService.CalculateMaximumSingleEmissions(model);

        return Ok(result);
    }
    
    [HttpPost("calculate/vehicle-flow")]
    public IActionResult CalculateVehicleFlowEmissions([FromBody] VehicleFlowEmissionsCalculateModel model)
    {
        var result = _emissionService.CalculateVehicleFlowEmissionsBatch(model);

        return Ok(result);
    }
    
    [HttpPost("calculate/traffic-light-queue")]
    public IActionResult CalculateTrafficLightQueueEmissions([FromBody] TrafficLightQueueEmissionsCalculateModel model)
    {
        var result = _emissionService.CalculateTrafficLightQueueEmissionsBatch(model);

        return Ok(result);
    }

    [HttpPost("calculate/open-coal-warehouse")]
    public IActionResult CalculateOpenCoalWarehouseEmissions([FromBody] OpenCoalWarehouseEmissionsCalculateModel model)
    {
        var result = _emissionService.CalculateOpenCoalWarehouseEmissions(model);

        return Ok(result);
    }
}