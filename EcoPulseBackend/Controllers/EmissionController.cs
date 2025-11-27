using EcoPulseBackend.Enums;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.Calculate;
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
    [HttpPost("calculate/gasoline-generator")]
    public IActionResult CalculateGasolineGenerator([FromBody] GasolineGeneratorEmissionsCalculateModel model)
    {
        var result = _emissionService.CalculateGasolineGeneratorEmissionsBatch(
            new List<Pollutant> { Pollutant.CO, Pollutant.CH, Pollutant.NO2, Pollutant.NO, Pollutant.SO2 },
            model.WorkHoursPerDay, model.WorkDaysPerYear,
            model.GeneratorCount, model.SameGeneratorCount);

        return Ok(result);
    }
    
    [HttpPost("calculate/reservoirs")]
    public IActionResult CalculateReservoirs([FromBody] ReservoirsEmissionsCalculateModel model)
    {
        var vaporConcentration = DataStorage.VaporConcentration[model.ReservoirType][model.ClimateZone][model.OilProduct];
        var result = _emissionService.CalculateReservoirsEmissionsBatch(
            new List<Pollutant> { Pollutant.RPK240280, Pollutant.H2S }, vaporConcentration,
            model.AutumnWinterOilAmount, model.SpringSummerOilAmount,
            model.DrainedVolume, model.AverageDrainTime);

        return Ok(result.Emissions);
    }
    
    [HttpPost("calculate/during-metal-machining")]
    public IActionResult CalculateDuringMetalMachining([FromBody] DuringMetalMachiningEmissionsCalculateModel model)
    { 
        var result = _emissionService.CalculateDuringMetalMachiningEmissions(model.MetalMachiningMachineType, model.WorkDaysPerYear);

        return Ok(result);
    }
    
    [HttpPost("calculate/during-welding-operations")]
    public IActionResult CalculateDuringWeldingOperations([FromBody] DuringWeldingOperationsEmissionsCalculateModel model)
    {
        var result = _emissionService.CalculateDuringWeldingOperationsEmissionsBatch(
                new List<Pollutant> { Pollutant.Fe2O3, Pollutant.MnO2, Pollutant.FluorideGases },
                model.ElectrodesPerYear,
                model.WorkDaysPerYear);

        return Ok(result.Emissions);
    }
    
    [HttpPost("calculate/maximum-single")]
    public IActionResult CalculateMaximumSingle([FromBody] MaximumSingleEmissionsCalculateModel model)
    {
        var result = _emissionService.CalculateMaximumSingleEmissions(Pollutant.SP, model);

        return Ok(result);
    }
    
    [HttpPost("calculate/vehicle-flow")]
    public IActionResult CalculateVehicleFlowEmissions([FromBody] List<VehicleGroup> vehicleGroups, float length)
    {
        var result = _emissionService.CalculateVehicleFlowEmissionsBatch(new List<Pollutant>
        {
            Pollutant.CO, Pollutant.NO2, Pollutant.CH, Pollutant.Soot,
            Pollutant.SO2, Pollutant.LeadCompounds, Pollutant.CH2O, Pollutant.C20H12
        }, vehicleGroups, length);

        return Ok(result);
    }
    
    [HttpPost("calculate/traffic-light-queue")]
    public IActionResult CalculateTrafficLightQueueEmissions([FromBody] List<VehicleGroupQueue> vehicleGroups, int trafficLightCycles, float trafficLightStopTime)
    {
        var result = _emissionService.CalculateTrafficLightQueueEmissionsBatch(new List<Pollutant>
        {
            Pollutant.CO, Pollutant.NO2, Pollutant.CH, Pollutant.Soot,
            Pollutant.SO2, Pollutant.LeadCompounds, Pollutant.CH2O, Pollutant.C20H12
        }, vehicleGroups, trafficLightCycles, trafficLightStopTime);

        return Ok(result);
    }

    /*[HttpPost("reports/gasoline-generator")]
    public IActionResult GetGasolineGeneratorReport([FromBody] GasolineGeneratorEmissionsReport report)
    {
        report.Emissions = _emissionService.CalculateGasolineGeneratorEmissionsBatch(
            new List<Pollutant> { Pollutant.CO, Pollutant.CH, Pollutant.NO2, Pollutant.NO, Pollutant.SO2 },
            report.WorkHoursPerDay, report.WorkDaysPerYear,
            report.GeneratorCount, report.SameGeneratorCount);

        var fileName = $"ИЗА_{report.PollutionSource}_{report.SelectionSource} Бензогенератор.xlsx";
        var stream = _exportService.CreateGasolineGeneratorEmissionsReport(report);

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

        var fileName = $"ИЗА_{report.PollutionSource}_{report.SelectionSource} Резервуары.xlsx";
        var stream = _exportService.CreateReservoirsEmissionsReport(report);

        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
    
    [HttpPost("reports/during-metal-machining")]
    public IActionResult GetDuringMetalMachiningReport([FromBody] DuringMetalMachiningEmissionsReport report)
    {
        report.Result = _emissionService.CalculateDuringMetalMachiningEmissions(report.MetalMachiningMachineType, report.WorkDaysPerYear);

        var fileName = $"ИЗА_{report.PollutionSource}_{report.SelectionSource} {report.MetalMachiningMachineType.GetDescription()}.xlsx";
        var stream = _exportService.CreateDuringMetalMachiningEmissionsReport(report);

        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
    
    [HttpPost("reports/during-welding-operations")]
    public IActionResult GetDuringWeldingOperationsReport([FromBody] DuringWeldingOperationsEmissionsReport report)
    {
        report.Result = _emissionService.CalculateDuringWeldingOperationsEmissionsBatch(
                new List<Pollutant> { Pollutant.Fe2O3, Pollutant.MnO2, Pollutant.FluorideGases },
                report.ElectrodesPerYear,
                report.WorkDaysPerYear);

        var fileName = $"ИЗА_{report.PollutionSource}_{report.SelectionSource} Сварочный аппарат.xlsx";
        var stream = _exportService.CreateDuringWeldingOperationsEmissionsReport(report);

        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }*/
}