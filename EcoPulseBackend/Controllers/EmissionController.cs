using EcoPulseBackend.Contexts;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.Calculate;
using EcoPulseBackend.Models.DangerZone;
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
    public IActionResult CalculateMaximumSingleEmissions([FromBody] MaximumSingleEmissionsCalculateModel model)
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
    
    [HttpPost("calculate/maximum-single-danger-zone")]
    public IActionResult CalculateMaximumSingleEmissionsDangerZone([FromBody] MaximumSingleEmissionsCalculateModel model)
    {
        var result = _emissionService.CalculateMaximumSingleEmissionsDangerZone(model);

        return Ok(result);
    }
    
    [HttpPost("calculate/danger-zones")]
    public IActionResult CalculateDangerZones([FromBody] DangerZoneCalculateModel model)
    {
        var emissionSources = _dbContext.EmissionSources.ToList();

        var result = new List<DangerZoneParameters>();
        
        foreach (var emissionSource in emissionSources)
        {
            var calculateModel = new MaximumSingleEmissionsCalculateModel
            {
                Pollutant = model.Pollutant,
                EjectedTemp =  emissionSource.EjectedTemp,
                AirTemp = model.AirTemp,
                AvgExitSpeed = emissionSource.AvgExitSpeed,
                HeightSource = emissionSource.HeightSource,
                DiameterSource = emissionSource.DiameterSource,
                TempStratificationRatio = emissionSource.TempStratificationRatio,
                SedimentationRateRatio = emissionSource.SedimentationRateRatio,
                WindSpeed = model.WindSpeed,
                Distance = 10000
            };

            var dangerZoneParameters = _emissionService.CalculateMaximumSingleEmissionsDangerZone(calculateModel);
            dangerZoneParameters.Lon = emissionSource.Lon;
            dangerZoneParameters.Lat = emissionSource.Lat;
            dangerZoneParameters.Angle = model.WindDirection;
            
            result.Add(dangerZoneParameters);
        }

        return Ok(result);
    }
}