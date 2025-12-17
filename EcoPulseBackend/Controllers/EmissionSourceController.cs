using EcoPulseBackend.Contexts;
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

    [HttpGet("/emissionSource")]
    public IActionResult GetEmissionSourceList()
    {
        var result = _dbContext.EmissionSources.ToList();

        return Ok(result);
    } 
}