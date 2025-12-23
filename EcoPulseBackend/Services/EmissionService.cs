using EcoPulseBackend.Enums;
using EcoPulseBackend.Extensions;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DangerZone;
using EcoPulseBackend.Models.DuringMetalMachining;
using EcoPulseBackend.Models.DuringWeldingOperations;
using EcoPulseBackend.Models.GasolineGenerator;
using EcoPulseBackend.Models.MaximumSingle;
using EcoPulseBackend.Models.OpenCoalWarehouse;
using EcoPulseBackend.Models.Reservoirs;
using EcoPulseBackend.Models.TrafficLightQueue;
using EcoPulseBackend.Models.TrafficLightQueueEmissionSource;
using EcoPulseBackend.Models.VehicleFlow;
using EcoPulseBackend.Models.VehicleFlowEmissionSource;

namespace EcoPulseBackend.Services;

public class EmissionService : IEmissionService
{
    public List<EmissionsResult> CalculateGasolineGeneratorEmissionsBatch(GasolineGeneratorEmissionsCalculateModel model)
    {
        var pollutants = new List<Pollutant> { Pollutant.CO, Pollutant.CH, Pollutant.NO2, Pollutant.NO, Pollutant.SO2 };
        
        var result = new List<EmissionsResult>();

        foreach (var pollutant in pollutants.OrderBy(p => (int)p))
        {
            result.Add(CalculateGasolineGeneratorEmissions(pollutant, model));
        }

        return result;
    }

    public ReservoirsEmissionsBatchResult CalculateReservoirsEmissionsBatch(ReservoirsEmissionsCalculateModel model)
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

    public List<EmissionsResult> CalculateDuringMetalMachiningEmissionsBatch(DuringMetalMachiningEmissionsCalculateModel model)
    {
        var pollutants = new List<Pollutant> { Pollutant.Fe2O3 };

        return pollutants.OrderBy(p => (int)p).Select(pollutant => CalculateDuringMetalMachiningEmissions(pollutant, model)).ToList();
    }

    public DuringWeldingOperationsEmissionsBatchResult CalculateDuringWeldingOperationsEmissionsBatch(DuringWeldingOperationsEmissionsCalculateModel model)
    {
        var pollutants = new List<Pollutant> { Pollutant.Fe2O3, Pollutant.MnO2, Pollutant.FluorideGases };
        
        var normElectrodesPerYear = model.ElectrodesPerYear * (100f - 15f) * 1e-2f;
        var materialsConsumption = normElectrodesPerYear / model.WorkDaysPerYear;
        
        var result = new DuringWeldingOperationsEmissionsBatchResult
        {
            NormElectrodesPerYear = normElectrodesPerYear,
            MaterialsConsumption = materialsConsumption,
            Emissions = new List<EmissionsResult>()
        };

        foreach (var pollutant in pollutants.OrderBy(p => (int)p))
        {
            result.Emissions.Add(CalculateDuringWeldingOperationsEmissions(pollutant, model.WorkDaysPerYear, materialsConsumption));
        }

        return result;
    }
    
    public EmissionsGroupResult CalculateMaximumSingleEmissions(MaximumSingleEmissionsCalculateModel model)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == model.Pollutant);

        if (!pollutantInfo.Mass.HasValue)
        {
            return new EmissionsGroupResult();
        }
        
        Setup(model);
        
        var distances = Enumerable.Range(1, model.Distance / 5).Select(i => i * 5f).ToList();
        
        var concentrations = GetNormalSurfaceConcentration(distances, (float)pollutantInfo.Mass); 
        
        var result = new EmissionsGroupResult { PollutantInfo = pollutantInfo };
        
        var topConcentrations = concentrations
            .Select((c, i) => new { Value = c, Index = i })
            .OrderByDescending(x => x.Value)
            .Take(model.MaxCount)
            .OrderBy(x => distances[x.Index]) 
            .ToList();

        foreach (var item in topConcentrations)
        {
            result.Emissions.Add(new EmissionsResult
            {
                MaximumEmission = item.Value,
                Distance = distances[item.Index]
            });
        }

        return result;
    }

    public List<EmissionsResult> CalculateVehicleFlowEmissionsBatch(VehicleFlowEmissionsCalculateModel model)
    {
        var pollutants = new List<Pollutant>
        {
            Pollutant.CO, Pollutant.NO2, Pollutant.CH, Pollutant.Soot,
            Pollutant.SO2, Pollutant.LeadCompounds, Pollutant.CH2O, Pollutant.C20H12
        };

        return pollutants.OrderBy(p => (int)p).Select(pollutant => CalculateVehicleFlowEmissions(pollutant, model)).OfType<EmissionsResult>().ToList();
    }

    public List<EmissionsResult> CalculateTrafficLightQueueEmissionsBatch(TrafficLightQueueEmissionsCalculateModel model)
    {
        var pollutants = new List<Pollutant>
        {
            Pollutant.CO, Pollutant.NO2, Pollutant.CH, Pollutant.Soot,
            Pollutant.SO2, Pollutant.LeadCompounds, Pollutant.CH2O, Pollutant.C20H12
        };

        return pollutants.OrderBy(p => (int)p).Select(pollutant => CalculateTrafficLightQueueEmissions(pollutant, model)).OfType<EmissionsResult>().ToList();
    }

    public List<EmissionsResult> CalculateOpenCoalWarehouseEmissions(OpenCoalWarehouseEmissionsCalculateModel model)
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
    
    public SingleDangerZone CalculateMaximumSingleEmissionsDangerZone(MaximumSingleEmissionsCalculateModel model)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == model.Pollutant);

        if (!pollutantInfo.Mass.HasValue)
        {
            return new SingleDangerZone();
        }
        
        Setup(model);
        
        var distances = Enumerable.Range(1, model.Distance / 5).Select(i => i * 5f).ToList();
        
        var concentrations = GetNormalSurfaceConcentration(distances, (float)pollutantInfo.Mass);
        return CalculateSingleDangerZone(concentrations, (float)pollutantInfo.Mass, model.WindSpeed);
    }

    public List<VehicleFlowDangerZone> CalculateVehicleFlowEmissionDangerZones(List<VehicleFlowEmissionSource> emissionSources)
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

    public List<TrafficLightQueueDangerZone> CalculateTrafficLightQueueEmissionDangerZones(List<TrafficLightQueueEmissionSource> emissionSources)
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

    #region Private

    private static EmissionsResult CalculateGasolineGeneratorEmissions(Pollutant pollutant, GasolineGeneratorEmissionsCalculateModel model)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == pollutant);

        if (!pollutantInfo.SpecificEmission.HasValue)
        {
            return new EmissionsResult();
        }
        
        var maximumEmission = 0.25f * (float)pollutantInfo.SpecificEmission * 5f * model.SameGeneratorCount / 3600f;
        var grossEmission = 0.25f * (float)pollutantInfo.SpecificEmission * 5f * model.WorkHoursPerDay * model.WorkDaysPerYear *
                            model.GeneratorCount * 1e-6f;

        var result = new EmissionsResult
        {
            PollutantInfo = pollutantInfo,
            MaximumEmission = maximumEmission,
            GrossEmission = grossEmission
        };

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
    
    private static EmissionsResult CalculateDuringWeldingOperationsEmissions(Pollutant pollutant, int workDaysPerYear, float materialsConsumption)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == pollutant);
        pollutantInfo.SpecificEmission = DataStorage.SpecificEmissionsByElectrodes.GetValueOrDefault(pollutant, 0f);
        
        if (!pollutantInfo.SpecificEmission.HasValue)
        {
            return new EmissionsResult();
        }
        
        var maximumEmission = materialsConsumption * (float)pollutantInfo.SpecificEmission * 0.4f / 3600;
        var grossEmission = maximumEmission * workDaysPerYear * 3.6f * 1e-3f;
        
        var result = new EmissionsResult
        {
            PollutantInfo = pollutantInfo,
            MaximumEmission = maximumEmission,
            GrossEmission = grossEmission
        };
        
        return result;
    }

    private float _heightSource; 
    private int _sedimentationRateRatio; 
    private float _diameterSource; 
    private int _tempStratificationRatio; 
    private float _avgExitSpeed; 
    private float _ejectedTemp; 
    private float _airTemp;

    private float _tempDiff;
    private float _volumeFlow;
    private float _riseVelocity;
    private float _velocityRatio;
    private float _buoyancyParam;
    private float _effectiveBuoyancy;
    
    private void Setup(MaximumSingleEmissionsCalculateModel model)
    {
        _heightSource = model.HeightSource;
        _sedimentationRateRatio = (int)model.SedimentationRateRatio;
        _diameterSource = model.DiameterSource;
        _tempStratificationRatio = (int)model.TempStratificationRatio;
        _avgExitSpeed = model.AvgExitSpeed;
        _ejectedTemp = model.EjectedTemp;
        _airTemp = model.AirTemp;

        _tempDiff = _ejectedTemp - _airTemp;
        _volumeFlow = (float)Math.PI * (float)Math.Pow(_diameterSource, 2) / 4f * _avgExitSpeed; 
        _riseVelocity = 0.65f * (float)Math.Pow(_volumeFlow * _tempDiff / _heightSource, 1f / 3f);
        _velocityRatio = 1.3f * _avgExitSpeed * _diameterSource / _heightSource;
        _buoyancyParam = 1000f * ((float)Math.Pow(_avgExitSpeed, 2) * _diameterSource) / ((float)Math.Pow(_heightSource, 2) * _tempDiff);
        _effectiveBuoyancy = 800f * (float)Math.Pow(_velocityRatio, 3);
    }

    private List<float> GetNormalSurfaceConcentration(List<float> distances, float mass)
    {
        var concentrations = new List<float>();

        float s1 = 0;
        var maxConcentration = GetMaximumSingleSurfaceConcentration(mass);

        var maxDistance = GetDistanceFromEmissionSourceSingle();

        foreach (var distance in distances)
        {
            var xDiv = distance / maxDistance;

            switch (xDiv)
            {
                case <= 1:
                    s1 = 3f * (float)Math.Pow(xDiv, 4) - 8f * (float)Math.Pow(xDiv, 3) + 6f * (float)Math.Pow(xDiv, 2);
                    break;
                case <= 8:
                    s1 = 1.13f / (0.13f * (float)Math.Pow(xDiv, 2) + 1f);
                    break;
                case <= 100f when _sedimentationRateRatio <= 1.5f:
                    s1 = xDiv / (3.556f * (float)Math.Pow(xDiv, 2) - 35.2f * xDiv + 120f);
                    break;
                case <= 100f:
                    s1 = 1f / (0.1f * (float)Math.Pow(xDiv, 2) + 2.456f * xDiv - 17.8f);
                    break;
                case > 100f when _sedimentationRateRatio <= 1.5f:
                    s1 = 144.3f * (float)Math.Pow(xDiv, -7f / 3f);
                    break;
                case > 100f:
                    s1 = 37.76f * (float)Math.Pow(xDiv, -7f / 3f);
                    break;
            }

            if (_heightSource <= 10f && xDiv < 1f)
            {
                var s1H = 0.125f * (10f - _heightSource) + 0.125f * (_heightSource - 2f) * s1;
                concentrations.Add(s1H * maxConcentration);
                return concentrations;
            }

            concentrations.Add(s1 * maxConcentration);
        }

        return concentrations;
    }

    private float GetMaximumSingleSurfaceConcentration(float mass)
    {
        float cM;

        const float nu = 1; //GetReliefCorrectionFactor();

        float m = 0;
        float n = 0;
        if (_buoyancyParam < 100f)
        {
            m = 1f / (0.67f + 0.1f * (float)Math.Sqrt(_buoyancyParam) + 0.34f * (float)Math.Pow(_buoyancyParam, 1f / 3f));
            
            if (_riseVelocity < 2f)
            {
                n = 0.532f * (float)Math.Pow(_riseVelocity, 2) - 2.13f * _riseVelocity + 3.13f;
            }
            else if (_riseVelocity < 0.5f)
            {
                n = 4.4f * _riseVelocity;
                var mS = 2.86f * m;
                cM = _tempStratificationRatio * mass * _sedimentationRateRatio * mS * nu / (float)Math.Pow(_heightSource, 7f / 3f);
                return cM;
            }
            else
            {
                n = 1;
            }
        }
        else if (_buoyancyParam >= 100f || _tempDiff is >= 0f and < 0.5f)
        {
            if (_buoyancyParam >= 100)
            {
                m = 1.47f / (float)Math.Pow(_buoyancyParam, 1f / 3f);
            }

            if (_velocityRatio >= 0.5f)
            {
                var k = _diameterSource / 8f * _volumeFlow;
                k = 1f / 7.1f * (float)Math.Sqrt(_avgExitSpeed * _volumeFlow);
                cM = _tempStratificationRatio * mass * _sedimentationRateRatio * n * nu * k / (float)Math.Pow(_heightSource, 4f / 3f);
                return cM;
            }
            else
            {
                float mS = 0.9f;
                cM = _tempStratificationRatio * mass * _sedimentationRateRatio * mS * nu / (float)Math.Pow(_heightSource, 7f / 3f);
                return cM;
            }
        }


        cM = (_tempStratificationRatio * mass * _sedimentationRateRatio * m * n * nu / ((float)Math.Pow(_heightSource, 2) * (float)Math.Pow(_volumeFlow * _tempDiff, 1f / 3f)));
        return cM;
    }

    private float GetDistanceFromEmissionSourceSingle()
    {
        float maxDistance = 0;

        float d = 0;
        if (_buoyancyParam < 100f)
        {
            if (_riseVelocity <= 0.5f)
            {
                d = 2.48f * (1f + 0.28f * (float)Math.Pow(_effectiveBuoyancy, 1f / 3f));
            }
            else if (_riseVelocity <= 2f)
            {
                d = 4.95f * _riseVelocity * (1f + 0.28f * (float)Math.Pow(_buoyancyParam, 1f / 3f));
            }
            else
            {
                d = 7f * (float)Math.Sqrt(_riseVelocity) * (1f + 0.28f * (float)Math.Pow(_buoyancyParam, 1f / 3f));
            }
        }
        else if (_buoyancyParam >= 100f || _tempDiff is >= 0f and < 0.5f)
        {
            if (_velocityRatio <= 0.5f)
            {
                d = 5.7f;
            }
            else if (_velocityRatio <= 2f)
            {
                d = 11.4f * _velocityRatio;
            }
            else
            {
                d = 16f * (float)Math.Sqrt(_velocityRatio);
            }
        }

        if (_velocityRatio is >= 0f and < 0.5f && _tempDiff is >= -0.5f and <= 0)
        {
            maxDistance = 5.7f * _heightSource;
            return maxDistance;
        }

        maxDistance = ((5f - _sedimentationRateRatio) / 4f) * d * _heightSource;
        return maxDistance;
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
    
    private EmissionsResult? CalculateTrafficLightQueueEmissions(Pollutant pollutant, TrafficLightQueueEmissionsCalculateModel model)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == pollutant);
        
        if (!DataStorage.VehicleEmissionFactors[model.VehicleGroups.First().VehicleType].ContainsKey(pollutant))
        {
            return null;
        }
        
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
    
    private SingleDangerZone CalculateSingleDangerZone(List<float> concentrations, float mass, float windSpeed)
    {
        var maxIndex = int.MinValue;
        var maxDistance = double.MinValue;
        var minDistance = double.MinValue;

        var valuesUpMax = new List<double>();
        
        var maxConcentration = concentrations.Max();
        for (var i = 0; i < concentrations.Count; i++)
        {
            valuesUpMax.Add(concentrations[i]);

            if (maxConcentration == concentrations[i])
            {
                maxIndex = i;
                maxDistance = (i + 1) * 5;
                break;
            }
        }

        double med;

        valuesUpMax.Sort();
        
        if (valuesUpMax.Count % 2 == 0)
        {
            med = (valuesUpMax[(valuesUpMax.Count / 2)] + valuesUpMax[(valuesUpMax.Count / 2) - 1]) / 2;
        }
        else
        {
            med = valuesUpMax[(valuesUpMax.Count / 2)];
        }

        for (var i = maxIndex; i < concentrations.Count; i++)
        {
            if (concentrations[i] < med)
            {
                minDistance = (i + 1) * 5;
                break;
            }
        }

        const int windAverageSpeed = 3;

        var windSpeedCoeff =  windSpeed != 0  ? windAverageSpeed / windSpeed : windAverageSpeed;
        
        var dangerZoneLength = (minDistance / Math.Sqrt(Math.Sqrt(mass))) * (1 / windSpeedCoeff);
        var dangerZoneWidth = Math.Round((minDistance - maxDistance) * 2 * windSpeedCoeff, 2) * Math.Sqrt(mass);

        var sortedConcentrations = concentrations.OrderByDescending(c => c).Take(100).ToList();
        var avgConcentration = sortedConcentrations.Average();
        
        var pm = avgConcentration * 1000;

        var color = DataStorage.ColorMap[225.4];
        foreach (var pair in DataStorage.ColorMap)
        {
            if (pm <= pair.Key)
            {
                color = pair.Value;
                break;
            }
        }
        
        return new SingleDangerZone()
        {
            Length = dangerZoneLength,
            Width = dangerZoneWidth,
            Color = color,
            AverageConcentration = avgConcentration
        };
    }

    #endregion
}