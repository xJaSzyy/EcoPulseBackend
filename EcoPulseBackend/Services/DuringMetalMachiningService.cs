using EcoPulseBackend.Enums;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DuringMetalMachining;

namespace EcoPulseBackend.Services;

public class DuringMetalMachiningService : IDuringMetalMachiningService
{
    public List<EmissionsResult> CalculateEmissionsBatch(DuringMetalMachiningEmissionsCalculateModel model)
    {
        var pollutants = new List<Pollutant> { Pollutant.Fe2O3 };

        return pollutants.OrderBy(p => (int)p).Select(pollutant => CalculateDuringMetalMachiningEmissions(pollutant, model)).ToList();
    }
    
    private static EmissionsResult CalculateDuringMetalMachiningEmissions(Pollutant pollutant, DuringMetalMachiningEmissionsCalculateModel model)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == pollutant);
        pollutantInfo.SpecificEmission = DataStorage.SpecificDustEmissionsByType.GetValueOrDefault(model.MetalMachiningMachineType, 0f);

        if (!pollutantInfo.SpecificEmission.HasValue)
        {
            return new EmissionsResult();
        }
        
        var maximumEmission = 0.2f * (float)pollutantInfo.SpecificEmission;
        var grossEmission = 0.2f * 3.6f * (float)pollutantInfo.SpecificEmission * model.WorkDaysPerYear * 1e-3f;

        var result = new EmissionsResult
        {
            PollutantInfo = pollutantInfo,
            MaximumEmission = maximumEmission,
            GrossEmission = grossEmission
        };
        
        return result;
    }
}