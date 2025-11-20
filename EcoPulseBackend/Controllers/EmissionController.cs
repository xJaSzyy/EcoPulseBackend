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

    [HttpPost("reports/gasoline-generator")]
    public IActionResult GetGasolineGeneratorReport([FromBody] GasolineGeneratorEmissionsReport report)
    {
        report.Emissions = _emissionService.CalculateGasolineGeneratorEmissionsBatch(
            new List<Pollutant> { Pollutant.CO, Pollutant.CH, Pollutant.NO2, Pollutant.NO, Pollutant.SO2 },
            report.WorkHoursPerDay, report.WorkDaysPerYear,
            report.GeneratorCount, report.SameGeneratorCount);

        var fileName = $"ИЗА_{report.PollutionSource}_{report.SelectionSource}.xlsx";
        var stream = _exportService.CreateGasolineGeneratorEmissionsReport(report, fileName);

        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
    
    [HttpPost("reports/reservoirs")]
    public IActionResult GetReservoirsReport([FromBody] ReservoirsEmissionsReport report)
    {
        report.VaporConcentration = DataStorage.VaporConcentration[report.ReservoirType][report.ClimateZone][report.OilProduct];
        report.Result = _emissionService.CalculateReservoirsEmissionsBatch(
            new List<Pollutant> { Pollutant.RPK240280, Pollutant.H2S }, report.VaporConcentration,
            report.AutumnWinterOilAmount, report.SpringSummerOilAmount,
            report.DrainedVolume, report.AverageDrainTime);

        var fileName = $"ИЗА_{report.PollutionSource}_{report.SelectionSource}.xlsx";
        var stream = _exportService.CreateReservoirsEmissionsReport(report, fileName);

        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
    
    [HttpPost("reports/during-metal-machining")]
    public IActionResult GetDuringMetalMachiningReport([FromBody] DuringMetalMachiningEmissionsReport report)
    {
        report.Result = _emissionService.CalculateDuringMetalMachiningEmissions(report.MetalMachiningMachineType, report.WorkDaysPerYear);

        var fileName = $"ИЗА_{report.PollutionSource}_{report.SelectionSource}.xlsx";
        var stream = _exportService.CreateDuringMetalMachiningEmissionsReport(report, fileName);

        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
    
    [HttpPost("reports/during-welding-operations")]
    public IActionResult GetDuringWeldingOperationsReport([FromBody] DuringWeldingOperationsEmissionsReport report)
    {
        report.Result = _emissionService.CalculateDuringWeldingOperationsEmissionsBatch(
                new List<Pollutant> { Pollutant.Fe2O3, Pollutant.MnO2, Pollutant.FluorideGases },
                report.ElectrodesPerYear,
                report.WorkDaysPerYear);

        var fileName = $"ИЗА_{report.PollutionSource}_{report.SelectionSource}.xlsx";
        var stream = _exportService.CreateDuringWeldingOperationsEmissionsReport(report, fileName);

        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
}