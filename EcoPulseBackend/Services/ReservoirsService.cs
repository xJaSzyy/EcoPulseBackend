using EcoPulseBackend.Enums;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.Reservoirs;

namespace EcoPulseBackend.Services;

public class ReservoirsService : IReservoirsService
{
    public ReservoirsEmissionsBatchResult CalculateEmissionsBatch(ReservoirsEmissionsCalculateModel model)
    {
        var pollutants = new List<Pollutant> { Pollutant.RPK240280, Pollutant.H2S };
        
        var vaporConcentration = DataStorage.VaporConcentration[model.ReservoirType][model.ClimateZone][model.OilProduct];
        
        var result = new ReservoirsEmissionsBatchResult
        {
            AnnualInjectionEmissions = (vaporConcentration.AutumnWinterVaporConcentration * model.AutumnWinterOilAmount + vaporConcentration.SpringSummerVaporConcentration * model.SpringSummerOilAmount) * 1e-6f,
            AnnualIrrigationEmissions = 50f * (model.AutumnWinterOilAmount + model.SpringSummerOilAmount) * 1e-6f,
            MaxVaporEmission = (vaporConcentration.MaxVaporConcentration * model.DrainedVolume) / model.AverageDrainTime,
            Emissions = new List<EmissionsResult>()
        };
        
        foreach (var pollutant in pollutants.OrderBy(p => (int)p))
        {
            result.Emissions.Add(CalculateReservoirsEmissions(pollutant, vaporConcentration, model));
        }
        
        return result;
    }
    
    private static EmissionsResult CalculateReservoirsEmissions(Pollutant pollutant, VaporConcentrationRecord vaporConcentration, ReservoirsEmissionsCalculateModel model)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == pollutant);

        if (!pollutantInfo.SpecificEmission.HasValue)
        {
            return new EmissionsResult();
        }
        
        var annualInjectionEmissions = (vaporConcentration.AutumnWinterVaporConcentration * model.AutumnWinterOilAmount +
                                        vaporConcentration.SpringSummerVaporConcentration * model.SpringSummerOilAmount) *
                                       1e-6f;
        var annualIrrigationEmissions = 50f * (model.AutumnWinterOilAmount + model.SpringSummerOilAmount) * 1e-6f;

        var maxVaporEmission = (vaporConcentration.MaxVaporConcentration * model.DrainedVolume) / model.AverageDrainTime;
        var grossEmission = annualInjectionEmissions + annualIrrigationEmissions;

        var result = new EmissionsResult
        {
            PollutantInfo = pollutantInfo,
            MaximumEmission = maxVaporEmission * (float)pollutantInfo.SpecificEmission * 1e-2f,
            GrossEmission = grossEmission * (float)pollutantInfo.SpecificEmission * 1e-2f,
        };

        return result;
    }
}