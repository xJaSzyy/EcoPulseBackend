using EcoPulseBackend.Enums;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DangerZone;
using EcoPulseBackend.Models.TrafficLightQueue;
using EcoPulseBackend.Models.TrafficLightQueueEmissionSource;

namespace EcoPulseBackend.Services;

public class TrafficLightQueueService : ITrafficLightQueueService
{
    public List<EmissionsResult> CalculateEmissionsBatch(TrafficLightQueueEmissionsCalculateModel model)
    {
        var pollutants = new List<Pollutant>
        {
            Pollutant.CO, Pollutant.NO2, Pollutant.CH, /*Pollutant.Soot,*/
            Pollutant.SO2, Pollutant.LeadCompounds, Pollutant.CH2O, Pollutant.C20H12
        };

        return pollutants.OrderBy(p => (int)p).Select(pollutant => CalculateTrafficLightQueueEmissions(pollutant, model)).OfType<EmissionsResult>().ToList();
    }
    
    public List<TrafficLightQueueDangerZone> CalculateDangerZones(List<TrafficLightQueueEmissionSource> emissionSources)
    {
        var result = new List<TrafficLightQueueDangerZone>();
        
        foreach (var source in emissionSources)
        {
            var calculateModel = new TrafficLightQueueEmissionsCalculateModel
            {
                VehicleGroups = source.VehicleGroups,
                TrafficLightCycles = source.TrafficLightCycles,
                TrafficLightStopTime = source.TrafficLightStopTime
            };
            
            var emissionsResult = CalculateTrafficLightQueueEmissions(Pollutant.NO2, calculateModel);

            if (emissionsResult == null)
            {
                return [];
            }
            
            var pm = emissionsResult.MaximumEmission;

            var color = DataStorage.ColorMap[225.4];
            foreach (var pair in DataStorage.ColorMap.Where(pair => pm <= pair.Key))
            {
                color = pair.Value;
                break;
            }
            
            result.Add(new TrafficLightQueueDangerZone
            {
                EmissionSourceId = source.Id,
                Location = source.Location,
                Color = color,
                AverageConcentration = emissionsResult.MaximumEmission
            });
        }

        return result;
    }
    
    private EmissionsResult? CalculateTrafficLightQueueEmissions(Pollutant pollutant, TrafficLightQueueEmissionsCalculateModel model)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == pollutant);
        
        var emission = 0f;

        foreach (var vehicleGroup in model.VehicleGroups)
        {
            var specificEmission = DataStorage.VehicleSpecificEmissions[vehicleGroup.VehicleType][pollutant];

            emission += specificEmission * vehicleGroup.VehiclesCount;
        }
        
        emission *= model.TrafficLightCycles * model.TrafficLightStopTime / 40f;

        var result = new EmissionsResult
        {
            MaximumEmission = emission,
            PollutantInfo =  pollutantInfo
        };
        
        return result;
    }
}