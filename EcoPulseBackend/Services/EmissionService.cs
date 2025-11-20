using EcoPulseBackend.Enums;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;

namespace EcoPulseBackend.Services;

public class EmissionService : IEmissionService
{
    public List<EmissionsResult> CalculateGasolineGeneratorEmissionsBatch(List<Pollutant> pollutants,
        int workHoursPerDay, int workDaysPerYear, int generatorCount, int sameGeneratorCount)
    {
        var result = new List<EmissionsResult>();

        foreach (var pollutant in pollutants.OrderBy(p => (int)p))
        {
            result.Add(CalculateGasolineGeneratorEmissions(pollutant, workHoursPerDay, workDaysPerYear, generatorCount, sameGeneratorCount));
        }

        return result;
    }

    public ReservoirsEmissionsBatchResult CalculateReservoirsEmissionsBatch(List<Pollutant> pollutants,
        VaporConcentrationRecord vaporConcentration, float autumnWinterOilAmount, float springSummerOilAmount,
        float drainedVolume, float averageDrainTime = 1200f)
    {
        var result = new ReservoirsEmissionsBatchResult
        {
            AnnualInjectionEmissions = (vaporConcentration.AutumnWinterVaporConcentration * autumnWinterOilAmount + vaporConcentration.SpringSummerVaporConcentration * springSummerOilAmount) * 1e-6f,
            AnnualIrrigationEmissions = 50f * (autumnWinterOilAmount + springSummerOilAmount) * 1e-6f,
            MaxVaporEmission = (vaporConcentration.MaxVaporConcentration * drainedVolume) / averageDrainTime,
            Emissions = new List<EmissionsResult>()
        };
        
        foreach (var pollutant in pollutants.OrderBy(p => (int)p))
        {
            result.Emissions.Add(CalculateReservoirsEmissions(pollutant, vaporConcentration, autumnWinterOilAmount, springSummerOilAmount, drainedVolume, averageDrainTime));
        }
        
        return result;
    }

    public EmissionsResult CalculateDuringMetalMachiningEmissions(MetalMachiningMachineType metalMachiningMachineType, int workDaysPerYear)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == Pollutant.Fe2O3);
        pollutantInfo.SpecificEmission = DataStorage.SpecificDustEmissionsByType.GetValueOrDefault(metalMachiningMachineType, 0f);

        var maximumEmission = 0.2f * pollutantInfo.SpecificEmission;
        var grossEmission = 0.2f * 3.6f * pollutantInfo.SpecificEmission * workDaysPerYear * 1e-3f;

        var result = new EmissionsResult
        {
            PollutantInfo = pollutantInfo,
            MaximumEmission = maximumEmission,
            GrossEmission = grossEmission
        };
        
        return result;
    }

    public DuringWeldingOperationsEmissionsBatchResult CalculateDuringWeldingOperationsEmissionsBatch(List<Pollutant> pollutants,
        float electrodesPerYear, int workDaysPerYear)
    {
        var normElectrodesPerYear = electrodesPerYear * (100 - 15) * 1e-2f;
        var materialsConsumption = normElectrodesPerYear / workDaysPerYear;
        
        var result = new DuringWeldingOperationsEmissionsBatchResult
        {
            NormElectrodesPerYear = normElectrodesPerYear,
            MaterialsConsumption = materialsConsumption,
            Emissions = new List<EmissionsResult>()
        };

        foreach (var pollutant in pollutants.OrderBy(p => (int)p))
        {
            result.Emissions.Add(CalculateDuringWeldingOperationsEmissions(pollutant, workDaysPerYear, materialsConsumption));
        }

        return result;
    }

    private static EmissionsResult CalculateGasolineGeneratorEmissions(Pollutant pollutant,
        int workHoursPerDay, int workDaysPerYear, int generatorCount, int sameGeneratorCount)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == pollutant);

        var maximumEmission = 0.25f * pollutantInfo.SpecificEmission * 5f * sameGeneratorCount / 3600f;
        var grossEmission = 0.25f * pollutantInfo.SpecificEmission * 5f * workHoursPerDay * workDaysPerYear *
                            generatorCount * 1e-6f;

        var result = new EmissionsResult
        {
            PollutantInfo = pollutantInfo,
            MaximumEmission = maximumEmission,
            GrossEmission = grossEmission
        };

        return result;
    }
    
    private static EmissionsResult CalculateReservoirsEmissions(Pollutant pollutant,
        VaporConcentrationRecord vaporConcentration, float autumnWinterOilAmount, float springSummerOilAmount,
        float drainedVolume, float averageDrainTime = 1200f)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == pollutant);

        var annualInjectionEmissions = (vaporConcentration.AutumnWinterVaporConcentration * autumnWinterOilAmount +
                                        vaporConcentration.SpringSummerVaporConcentration * springSummerOilAmount) *
                                       1e-6f;
        var annualIrrigationEmissions = 50f * (autumnWinterOilAmount + springSummerOilAmount) * 1e-6f;

        var maxVaporEmission = (vaporConcentration.MaxVaporConcentration * drainedVolume) / averageDrainTime;
        var grossEmission = annualInjectionEmissions + annualIrrigationEmissions;

        var result = new EmissionsResult
        {
            PollutantInfo = pollutantInfo,
            MaximumEmission = maxVaporEmission * pollutantInfo.SpecificEmission * 1e-2f,
            GrossEmission = grossEmission * pollutantInfo.SpecificEmission * 1e-2f,
        };

        return result;
    }
    
    private static EmissionsResult CalculateDuringWeldingOperationsEmissions(Pollutant pollutant, int workDaysPerYear, float materialsConsumption)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == pollutant);
        pollutantInfo.SpecificEmission = DataStorage.SpecificEmissionsByElectrodes.GetValueOrDefault(pollutant, 0f);
        
        var maximumEmission = materialsConsumption * pollutantInfo.SpecificEmission * 0.4f / 3600;
        var grossEmission = maximumEmission * workDaysPerYear * 3.6f * 1e-3f;
        
        var result = new EmissionsResult
        {
            PollutantInfo = pollutantInfo,
            MaximumEmission = maximumEmission,
            GrossEmission = grossEmission
        };
        
        return result;
    }
}