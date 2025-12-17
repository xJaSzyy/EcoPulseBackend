using EcoPulseBackend.Contexts;
using EcoPulseBackend.Models.EmissionSource;
using Microsoft.AspNetCore.Mvc;

namespace EcoPulseBackend.Controllers;

[ApiController]
public class EmissionSourceController : ControllerBase
{
    private readonly ApplicationDbContext _dbContext;

    public EmissionSourceController(ApplicationDbContext dbContext)
    {
        _dbContext = dbContext;
    }

    [HttpPost("/emissionSource")]
    public async Task<IActionResult> AddEmissionSource([FromBody] EmissionSourceAddModel model)
    {
        var emissionSource = new EmissionSource
        {
            Lon = model.Lon,
            Lat = model.Lat,
            EjectedTemp = model.EjectedTemp,
            AvgExitSpeed = model.AvgExitSpeed,
            HeightSource = model.HeightSource,
            DiameterSource = model.DiameterSource,
            TempStratificationRatio = model.TempStratificationRatio,
            SedimentationRateRatio = model.SedimentationRateRatio
        };
        
        _dbContext.EmissionSources.Add(emissionSource);
        await _dbContext.SaveChangesAsync();
        
        return Ok(emissionSource);
    }

    [HttpGet("/emissionSource")]
    public IActionResult GetAllEmissionSources()
    {
        var result = _dbContext.EmissionSources.ToList();

        return Ok(result);
    } 
    
    [HttpPut("/emissionSource")]
    public async Task<IActionResult> UpdateEmissionSource([FromBody] EmissionSourceUpdateModel model)
    {
        var emissionSource = _dbContext.EmissionSources.FirstOrDefault(s => s.Id == model.Id);

        if (emissionSource == null)
        {
            return NotFound();
        }
        
        emissionSource.Lon = model.Lon;
        emissionSource.Lat = model.Lat;
        emissionSource.EjectedTemp = model.EjectedTemp;
        emissionSource.AvgExitSpeed = model.AvgExitSpeed;
        emissionSource.HeightSource = model.HeightSource;
        emissionSource.DiameterSource = model.DiameterSource;
        emissionSource.TempStratificationRatio = model.TempStratificationRatio;
        emissionSource.SedimentationRateRatio = model.SedimentationRateRatio;
        
        _dbContext.EmissionSources.Update(emissionSource);
        await _dbContext.SaveChangesAsync();
        
        return Ok(emissionSource);
    }
    
    [HttpDelete("/emissionSource/{id:int}")]
    public async Task<IActionResult> DeleteEmissionSource(int id)
    {
        var emissionSource = _dbContext.EmissionSources.FirstOrDefault(s => s.Id == id);

        if (emissionSource == null)
        {
            return NotFound();
        }
        
        _dbContext.EmissionSources.Remove(emissionSource);
        await _dbContext.SaveChangesAsync();
        
        return Ok(emissionSource);
    }
}