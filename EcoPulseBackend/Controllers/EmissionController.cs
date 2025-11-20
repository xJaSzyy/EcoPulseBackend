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

    [HttpPost("report/gasolineGenerator")]
    public IActionResult GetGasolineGeneratorReport([FromBody] GasolineGeneratorEmissionsReport report)
    {
        report.Emissions = _emissionService.CalculateGasolineGeneratorEmissionsBatch(
            new List<Pollutant> { Pollutant.CO, Pollutant.CH, Pollutant.NO2, Pollutant.NO, Pollutant.SO2 },
            report.WorkHoursPerDay, report.WorkDaysPerYear,
            report.GeneratorCount, report.SameGeneratorCount);

        var stream = _exportService.CreateGasolineGeneratorEmissionsReport(report);
        var fileName = $"ИЗА_{report.PollutionSource}_{report.SelectionSource}.xlsx";

        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
    
    [HttpPost("report/reservoirs")]
    public IActionResult GetReservoirsReport([FromBody] ReservoirsEmissionsReport report)
    {
        report.VaporConcentration = DataStorage.VaporConcentration[report.ReservoirType][report.ClimateZone][report.OilProduct];
        report.Result = _emissionService.CalculateReservoirsEmissionsBatch(
            new List<Pollutant> { Pollutant.RPK240280, Pollutant.H2S }, report.VaporConcentration,
            report.AutumnWinterOilAmount, report.SpringSummerOilAmount,
            report.DrainedVolume, report.AverageDrainTime);

        var stream = _exportService.CreateReservoirsEmissionsReport(report);
        var fileName = $"ИЗА_{report.PollutionSource}_{report.SelectionSource}.xlsx";

        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
}