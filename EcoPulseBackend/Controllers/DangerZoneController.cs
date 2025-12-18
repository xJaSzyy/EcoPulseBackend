using EcoPulseBackend.Contexts;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models.DangerZone;
using EcoPulseBackend.Models.MaximumSingle;
using Microsoft.AspNetCore.Mvc;

namespace EcoPulseBackend.Controllers;

[ApiController]
public class DangerZoneController : ControllerBase
{
    private readonly IEmissionService _emissionService;
    private readonly ApplicationDbContext _dbContext;

    public DangerZoneController(IEmissionService emissionService, ApplicationDbContext dbContext)
    {
        _emissionService = emissionService;
        _dbContext = dbContext;
    }
    
    [HttpPost("danger-zone/single")]
    public IActionResult CalculateMaximumSingleEmissionsDangerZone([FromBody] MaximumSingleEmissionsCalculateModel model)
    {
        var result = _emissionService.CalculateMaximumSingleEmissionsDangerZone(model);
        
        return Ok(result);
    }
    
    [HttpPost("danger-zones/single")]
    public IActionResult CalculateSingleDangerZones([FromBody] DangerZoneCalculateModel model)
    {
        var emissionSources = _dbContext.SingleEmissionSources.ToList();

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
            dangerZoneParameters.EmissionSourceId = emissionSource.Id;
            dangerZoneParameters.Lon = emissionSource.Location.Lon;
            dangerZoneParameters.Lat = emissionSource.Location.Lat;
            dangerZoneParameters.Angle = model.WindDirection;
            
            result.Add(dangerZoneParameters);
        }

        return Ok(result);
    }
    
    [HttpPost("danger-zones/vehicleFlow")]
    public IActionResult CalculateVehicleFlowDangerZones()
    {
        var emissionSources = _dbContext.VehicleFlowEmissionSources.ToList();

        return Ok(emissionSources);
    }
}