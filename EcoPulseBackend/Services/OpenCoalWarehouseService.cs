using EcoPulseBackend.Enums;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.OpenCoalWarehouse;

namespace EcoPulseBackend.Services;

public class OpenCoalWarehouseService : IOpenCoalWarehouseService
{
    public List<EmissionsResult> CalculateEmissions(OpenCoalWarehouseEmissionsCalculateModel model)
    {
        // Расчет пылевыделения при перегрузочных работах (п. 6.3)									
        const float humidityFactor = 0.7f;
        const float averageWindSpeedFactor = 1.2f;
        const float pileHeightFactor = 2.5f;
        const float protectionDegreeFactor = 1f;
        const float maxWindSpeedFactor = 1.7f;

        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == Pollutant.CoalDust);

        var grossEmission1 = model.SpecificEmission * model.UnloadMaterialCountPerYear * humidityFactor * averageWindSpeedFactor *
                            pileHeightFactor * protectionDegreeFactor * 1e-6f * (1f - model.DustSuppressionEfficiency);
        var maximumEmission1 = model.SpecificEmission * model.UnloadMaterialCountPerHour * humidityFactor * maxWindSpeedFactor *
            pileHeightFactor * protectionDegreeFactor * (1f - model.DustSuppressionEfficiency) / 3600f;

        // Расчет количества пыли, сдуваемой с поверхности открытого угольного склада (п. 9)									
        const float specificWindErosionRate = 1e-6f;
        const float rockCrushingFactor = 0.1f;
        const float surfaceProfileFactor = 1.45f;

        var grossEmission2 = 86.4f * specificWindErosionRate * model.CoalPileBaseArea * rockCrushingFactor *
                             humidityFactor * averageWindSpeedFactor * protectionDegreeFactor * surfaceProfileFactor *
                             (365 - (model.SnowyDaysCount + model.RainyDaysCount)) *
                             (1 - model.DustSuppressionEfficiency);

        var maximumEmission2 = specificWindErosionRate * model.CoalPileBaseArea * humidityFactor * maxWindSpeedFactor *
                               protectionDegreeFactor * surfaceProfileFactor * rockCrushingFactor *
                               (1 - model.DustSuppressionEfficiency) * 1000f;
        
        var result = new EmissionsResult
        {
            PollutantInfo = pollutantInfo,
            MaximumEmission = maximumEmission1 + maximumEmission2,
            GrossEmission = grossEmission1 + grossEmission2
        };

        return new List<EmissionsResult> { result };
    }
}