using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models.Recommendation;
using Microsoft.AspNetCore.Mvc;

namespace EcoPulseBackend.Controllers;

public class RecommendationController : ControllerBase
{
    private readonly IRecommendationService _service;

    public RecommendationController(IRecommendationService service)
    {
        _service = service;
    }

    [HttpGet("recommendation")]
    public async Task<IActionResult> GetRecommendation(GetRecommendationModel model)
    {
        var result = await _service.GetRecommendation(model);
        
        return Ok(result);
    }
}