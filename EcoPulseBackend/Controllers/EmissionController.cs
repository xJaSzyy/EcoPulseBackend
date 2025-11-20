using EcoPulseBackend.Enums;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
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

    [HttpGet("report/gasolineGenerator")]
    public IActionResult GetGasolineGeneratorReport([FromQuery] string selectionSource, string pollutionSource,
        int workHoursPerDay, int workDaysPerYear, int generatorCount, int sameGeneratorCount)
    {
        var report = new GasolineGeneratorEmissionsReport
        {
            SelectionSource = selectionSource,
            PollutionSource = pollutionSource,
            WorkHoursPerDay = workHoursPerDay,
            WorkDaysPerYear = workDaysPerYear,
            GeneratorCount = generatorCount,
            SameGeneratorCount = sameGeneratorCount
        };
        report.Emissions = _emissionService.CalculateGasolineGeneratorEmissionsBatch(
            new List<Pollutant> { Pollutant.CO, Pollutant.CH, Pollutant.NO2, Pollutant.NO, Pollutant.SO2 },
            report.WorkHoursPerDay, report.WorkDaysPerYear,
            report.GeneratorCount, report.SameGeneratorCount);

        var stream = _exportService.CreateGasolineGeneratorEmissionsReport(report);
        var fileName = $"ИЗА_{report.PollutionSource}_{report.SelectionSource}.xlsx";

        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
}