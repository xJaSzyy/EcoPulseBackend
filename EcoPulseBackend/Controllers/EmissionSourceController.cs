using EcoPulseBackend.Contexts;
using EcoPulseBackend.Enums;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.SingleEmissionSource;
using EcoPulseBackend.Models.TrafficLightQueue;
using EcoPulseBackend.Models.TrafficLightQueueEmissionSource;
using EcoPulseBackend.Models.VehicleFlowEmissionSource;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace EcoPulseBackend.Controllers;

[ApiController]
public class EmissionSourceController : ControllerBase
{
    private readonly ApplicationDbContext _dbContext;

    public EmissionSourceController(ApplicationDbContext dbContext)
    {
        _dbContext = dbContext;
    }

    [HttpPost("/emission-source/single")]
    public async Task<IActionResult> AddSingleEmissionSource([FromBody] SingleEmissionSourceAddModel model)
    {
        var emissionSource = new SingleEmissionSource
        {
            Location = new Coordinates
            {
                Lon = model.Lon,
                Lat = model.Lat
            },
            EjectedTemp = model.EjectedTemp,
            AvgExitSpeed = model.AvgExitSpeed,
            HeightSource = model.HeightSource,
            DiameterSource = model.DiameterSource,
            TempStratificationRatio = model.TempStratificationRatio,
            SedimentationRateRatio = model.SedimentationRateRatio
        };
        
        _dbContext.SingleEmissionSources.Add(emissionSource);
        await _dbContext.SaveChangesAsync();
        
        return Ok(emissionSource);
    }
    
    [HttpGet("/emission-source/single/{id:int}")]
    public IActionResult GetSingleEmissionSourceById(int id)
    {
        var result = _dbContext.SingleEmissionSources.FirstOrDefault(s => s.Id == id);

        if (result == null)
        {
            return NotFound();
        }
        
        return Ok(result);
    } 
    
    [HttpPut("/emission-source/single")]
    public async Task<IActionResult> UpdateSingleEmissionSource([FromBody] SingleEmissionSourceUpdateModel model)
    {
        var emissionSource = _dbContext.SingleEmissionSources.FirstOrDefault(s => s.Id == model.Id);

        if (emissionSource == null)
        {
            return NotFound();
        }
        
        emissionSource.Location.Lon = model.Lon;
        emissionSource.Location.Lat = model.Lat;
        emissionSource.EjectedTemp = model.EjectedTemp;
        emissionSource.AvgExitSpeed = model.AvgExitSpeed;
        emissionSource.HeightSource = model.HeightSource;
        emissionSource.DiameterSource = model.DiameterSource;
        emissionSource.TempStratificationRatio = model.TempStratificationRatio;
        emissionSource.SedimentationRateRatio = model.SedimentationRateRatio;
        
        _dbContext.SingleEmissionSources.Update(emissionSource);
        await _dbContext.SaveChangesAsync();
        
        return Ok(emissionSource);
    }
    
    [HttpDelete("/emission-source/single/{id:int}")]
    public async Task<IActionResult> DeleteSingleEmissionSource(int id)
    {
        var emissionSource = _dbContext.SingleEmissionSources.FirstOrDefault(s => s.Id == id);

        if (emissionSource == null)
        {
            return NotFound();
        }
        
        _dbContext.SingleEmissionSources.Remove(emissionSource);
        await _dbContext.SaveChangesAsync();
        
        return Ok(emissionSource);
    }
    
    [HttpPost("/emission-source/vehicle-flow")]
    public async Task<IActionResult> AddVehicleFlowEmissionSource([FromBody] VehicleFlowEmissionSourceAddModel model)
    {
        var emissionSource = new VehicleFlowEmissionSource
        {
            StartLocation = model.StartLocation,
            EndLocation = model.EndLocation,    
            VehicleType = model.VehicleType,
            MaxTrafficIntensity = model.MaxTrafficIntensity,
            AverageSpeed =  model.AverageSpeed
        };
        
        _dbContext.VehicleFlowEmissionSources.Add(emissionSource);
        await _dbContext.SaveChangesAsync();
        
        return Ok(emissionSource);
    }
    
    [HttpGet("/emission-source/vehicle-flow/{id:int}")]
    public IActionResult GetVehicleFlowEmissionSourceById(int id)
    {
        var result = _dbContext.VehicleFlowEmissionSources.FirstOrDefault(s => s.Id == id);

        if (result == null)
        {
            return NotFound();
        }
        
        return Ok(result);
    }
    
    [HttpPut("/emission-source/vehicle-flow")]
    public async Task<IActionResult> UpdateVehicleFlowEmissionSource([FromBody] VehicleFlowEmissionSourceUpdateModel model)
    {
        var emissionSource = _dbContext.VehicleFlowEmissionSources.FirstOrDefault(s => s.Id == model.Id);

        if (emissionSource == null)
        {
            return NotFound();
        }

        emissionSource.StartLocation = model.StartLocation ??  emissionSource.StartLocation;
        emissionSource.EndLocation = model.EndLocation ??  emissionSource.EndLocation;
        emissionSource.VehicleType = model.VehicleType ?? emissionSource.VehicleType;
        emissionSource.MaxTrafficIntensity = model.MaxTrafficIntensity ??  emissionSource.MaxTrafficIntensity;
        emissionSource.AverageSpeed = model.AverageSpeed ?? emissionSource.AverageSpeed;
        
        _dbContext.VehicleFlowEmissionSources.Update(emissionSource);
        await _dbContext.SaveChangesAsync();
        
        return Ok(emissionSource);
    }
    
    [HttpPost("/emission-source/traffic-light-queue")]
    public async Task<IActionResult> AddTrafficLightQueueEmissionSource([FromBody] TrafficLightQueueEmissionSourceAddModel model)
    {
        var emissionSource = new TrafficLightQueueEmissionSource
        {
            Location = model.Location,
            TrafficLightCycles = model.TrafficLightCycles,
            TrafficLightStopTime = model.TrafficLightStopTime,
            VehicleGroups = model.VehicleGroups
        };
        
        _dbContext.TrafficLightQueueEmissionSources.Add(emissionSource);
        await _dbContext.SaveChangesAsync();
        
        return Ok(emissionSource);
    }
    
    [HttpGet("/emission-source/traffic-light-queue")]
    public IActionResult GetTrafficLightQueueEmissionSources()
    {
        var sources = _dbContext.TrafficLightQueueEmissionSources
            .Include(s => s.VehicleGroups)
            .ToList();

        return Ok(sources);
    }
}