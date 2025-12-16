using EcoPulseBackend.Enums;
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
        var emissionSources = new List<EmissionSource>
        {
            new()
            {
                Lon = 85.99424f,
                Lat = 55.347918f,
                EjectedTemp = 255,
                AvgExitSpeed = 30,
                HeightSource = 100,
                DiameterSource = 4,
                TempStratificationRatio = CoefficientRegion.BuryatiaOrTransBaikal,
                SedimentationRateRatio = CoefficientDegreePurification.Low,
            },
            new()
            {
                Lon = 86.068655f,
                Lat = 55.363112f,
                EjectedTemp = 245,
                AvgExitSpeed = 19,
                HeightSource = 120,
                DiameterSource = 3,
                TempStratificationRatio = CoefficientRegion.BuryatiaOrTransBaikal,
                SedimentationRateRatio = CoefficientDegreePurification.Low,
            },
            new()
            {
                Lon = 86.035864f,
                Lat = 55.365342f,
                EjectedTemp = 255,
                AvgExitSpeed = 15,
                HeightSource = 80,
                DiameterSource = 2,
                TempStratificationRatio = CoefficientRegion.BuryatiaOrTransBaikal,
                SedimentationRateRatio = CoefficientDegreePurification.Low,
            },
            new()
            {
                Lon = 86.076927f,
                Lat = 55.390792f,
                EjectedTemp = 265,
                AvgExitSpeed = 30,
                HeightSource = 60,
                DiameterSource = 6,
                TempStratificationRatio = CoefficientRegion.BuryatiaOrTransBaikal,
                SedimentationRateRatio = CoefficientDegreePurification.Low,
            }
        };

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