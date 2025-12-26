using EcoPulseBackend.Contexts;
using Microsoft.AspNetCore.Mvc;

namespace EcoPulseBackend.Controllers;

[ApiController]
public class CityController : ControllerBase
{
    private readonly ApplicationDbContext _dbContext;

    public CityController(ApplicationDbContext dbContext)
    {
        _dbContext = dbContext;
    }

    [HttpGet("city/{id:int}")]
    public IActionResult GetCityById(int id)
    {
        var city = _dbContext.Cities.FirstOrDefault(c => c.Id == id);

        if (city == null)
        {
            return NotFound();
        }

        return Ok(city);
    }
}