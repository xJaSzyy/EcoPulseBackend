using EcoPulseBackend.Enums;
using EcoPulseBackend.Extensions;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DangerZone;
using EcoPulseBackend.Models.VehicleFlow;
using EcoPulseBackend.Models.VehicleFlowEmissionSource;

namespace EcoPulseBackend.Services;

public class VehicleFlowService : IVehicleFlowService
{
    public List<EmissionsResult> CalculateEmissionsBatch(VehicleFlowEmissionsCalculateModel model)
    {
        var pollutants = new List<Pollutant>
        {
            Pollutant.CO, Pollutant.NO2, Pollutant.CH, Pollutant.Soot,
            Pollutant.SO2, Pollutant.LeadCompounds, Pollutant.CH2O, Pollutant.C20H12
        };

        return pollutants.OrderBy(p => (int)p).Select(pollutant => CalculateVehicleFlowEmissions(pollutant, model)).OfType<EmissionsResult>().ToList();
    }
    
    public List<VehicleFlowDangerZone> CalculateDangerZones(List<VehicleFlowEmissionSource> emissionSources)
    {
        var result = new List<VehicleFlowDangerZone>();
        
        foreach (var source in emissionSources)
        {
            var points = source.Points;
            
            float length = 0;
            for (var i = 1; i < points.Count; i++)
            {
                length += (float)GeoUtils.DistanceMeters(points[i - 1], points[i]);
            }
            
            var calculateModel = new VehicleFlowEmissionsCalculateModel
            {
                VehicleGroups =
                [
                    new VehicleGroup
                    {
                        VehicleType = source.VehicleType,
                        MaxTrafficIntensity = source.MaxTrafficIntensity * (length / 1000f),
                        AverageSpeed = source.AverageSpeed
                    }
                ],
                Length = length
            };
            
            var emissionsResult = CalculateVehicleFlowEmissions(Pollutant.NO2, calculateModel);

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
            
            result.Add(new VehicleFlowDangerZone
            {
                EmissionSourceId = source.Id,
                Points = points,
                Color = color,
                AverageConcentration = emissionsResult.MaximumEmission
            });
        }

        return result;
    }
    
    private static EmissionsResult? CalculateVehicleFlowEmissions(Pollutant pollutant, VehicleFlowEmissionsCalculateModel model)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == pollutant);
        
        if (!DataStorage.VehicleEmissionFactors[model.VehicleGroups.First().VehicleType].ContainsKey(pollutant))
        {
            return null;
        }
        
        var emission = 0f;

        foreach (var vehicleGroup in model.VehicleGroups)
        {
            var specificEmission = DataStorage.VehicleEmissionFactors[vehicleGroup.VehicleType][pollutant];

            var speedCorrectionFactor = DataStorage.GetSpeedCorrectionFactor(vehicleGroup.AverageSpeed);
            
            emission += specificEmission * vehicleGroup.MaxTrafficIntensity * speedCorrectionFactor;
        }
        
        emission *= model.Length / 3600f;

        var result = new EmissionsResult
        {
            MaximumEmission = emission,
            PollutantInfo =  pollutantInfo
        };
        
        return result;
    }
}