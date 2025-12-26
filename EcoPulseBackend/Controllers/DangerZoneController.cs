using EcoPulseBackend.Contexts;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models.DangerZone;
using EcoPulseBackend.Models.MaximumSingle;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

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
        var result = _emissionService.MaximumSingleService.CalculateDangerZone(model);
        
        return Ok(result);
    }
    
    [HttpPost("danger-zones/single")]
    public IActionResult CalculateSingleDangerZones([FromBody] SingleDangerZoneCalculateModel model)
    {
        var emissionSources = _dbContext.SingleEmissionSources.ToList();

        var result = new List<SingleDangerZone>();
        
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

            var dangerZone = _emissionService.MaximumSingleService.CalculateDangerZone(calculateModel);
            dangerZone.EmissionSourceId = emissionSource.Id;
            dangerZone.Lon = emissionSource.Location.Lon;
            dangerZone.Lat = emissionSource.Location.Lat;
            dangerZone.Angle = model.WindDirection;
            
            result.Add(dangerZone);
        }

        return Ok(result);
    }
    
    [HttpPost("danger-zones/vehicle-flow")]
    public IActionResult CalculateVehicleFlowDangerZones([FromBody] VehicleFlowDangerZoneCalculateModel model)
    {
        var emissionSources = _dbContext.VehicleFlowEmissionSources
            .Where(s => model.CityIds.Contains(s.CityId))
            .ToList();

        var result = _emissionService.VehicleFlowService.CalculateDangerZones(emissionSources);

        return Ok(result);
    }
    
    [HttpPost("danger-zones/traffic-light-queue")]
    public IActionResult CalculateTrafficLightQueueDangerZones()
    {
        var emissionSources = _dbContext.TrafficLightQueueEmissionSources
            .Include(s => s.VehicleGroups)
            .ToList();

        var result = _emissionService.TrafficLightQueueService.CalculateDangerZones(emissionSources);

        return Ok(result);
    }
}