using EcoPulseBackend.Enums;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.Calculate;

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

        if (!pollutantInfo.SpecificEmission.HasValue)
        {
            return new EmissionsResult();
        }
        
        var maximumEmission = 0.2f * (float)pollutantInfo.SpecificEmission;
        var grossEmission = 0.2f * 3.6f * (float)pollutantInfo.SpecificEmission * workDaysPerYear * 1e-3f;

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
    
    public List<EmissionsResult> CalculateMaximumSingleEmissions(Pollutant pollutant, MaximumSingleEmissionsCalculateModel model)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == pollutant);

        if (!pollutantInfo.Mass.HasValue)
        {
            return new List<EmissionsResult>();
        }
        
        Setup(model);
        
        var distances = Enumerable.Range(1, model.Distance / 5).Select(i => i * 5f).ToList();
        
        var concentrations = GetNormalSurfaceConcentration(distances, (float)pollutantInfo.Mass); 
        
        var result = new List<EmissionsResult>();
        
        foreach (var concentration in concentrations)
        {
            result.Add(new EmissionsResult
            {
                MaximumEmission = concentration
            });
        }

        return result;
    }

    public List<EmissionsResult> CalculateVehicleFlowEmissionsBatch(List<Pollutant> pollutants,
        List<VehicleGroup> vehicleGroups, float length)
    {
        var result = new List<EmissionsResult>();

        foreach (var pollutant in pollutants)
        {
            var emission = CalculateVehicleFlowEmissions(pollutant, vehicleGroups, length);
            if (emission != null)
            {
                result.Add(emission);
            }
        }

        return result;
    }

    public List<EmissionsResult> CalculateTrafficLightQueueEmissionsBatch(List<Pollutant> pollutants, List<VehicleGroupQueue> vehicleGroups, int trafficLightCycles, float trafficLightStopTime)
    {
        var result = new List<EmissionsResult>();

        foreach (var pollutant in pollutants)
        {
            var emission = CalculateTrafficLightQueueEmissions(pollutant, vehicleGroups, trafficLightCycles, trafficLightStopTime);
            if (emission != null)
            {
                result.Add(emission);
            }
        }

        return result;
    }

    #region Private

    private static EmissionsResult CalculateGasolineGeneratorEmissions(Pollutant pollutant,
        int workHoursPerDay, int workDaysPerYear, int generatorCount, int sameGeneratorCount)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == pollutant);

        if (!pollutantInfo.SpecificEmission.HasValue)
        {
            return new EmissionsResult();
        }
        
        var maximumEmission = 0.25f * (float)pollutantInfo.SpecificEmission * 5f * sameGeneratorCount / 3600f;
        var grossEmission = 0.25f * (float)pollutantInfo.SpecificEmission * 5f * workHoursPerDay * workDaysPerYear *
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

        if (!pollutantInfo.SpecificEmission.HasValue)
        {
            return new EmissionsResult();
        }
        
        var annualInjectionEmissions = (vaporConcentration.AutumnWinterVaporConcentration * autumnWinterOilAmount +
                                        vaporConcentration.SpringSummerVaporConcentration * springSummerOilAmount) *
                                       1e-6f;
        var annualIrrigationEmissions = 50f * (autumnWinterOilAmount + springSummerOilAmount) * 1e-6f;

        var maxVaporEmission = (vaporConcentration.MaxVaporConcentration * drainedVolume) / averageDrainTime;
        var grossEmission = annualInjectionEmissions + annualIrrigationEmissions;

        var result = new EmissionsResult
        {
            PollutantInfo = pollutantInfo,
            MaximumEmission = maxVaporEmission * (float)pollutantInfo.SpecificEmission * 1e-2f,
            GrossEmission = grossEmission * (float)pollutantInfo.SpecificEmission * 1e-2f,
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
    
    private static EmissionsResult? CalculateVehicleFlowEmissions(Pollutant pollutant, List<VehicleGroup> vehicleGroups, float length)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == pollutant);
        
        if (!DataStorage.VehicleEmissionFactors[vehicleGroups.First().VehicleType].ContainsKey(pollutant))
        {
            return null;
        }
        
        var emission = 0f;

        foreach (var vehicleGroup in vehicleGroups)
        {
            var specificEmission = DataStorage.VehicleEmissionFactors[vehicleGroup.VehicleType][pollutant];

            var speedCorrectionFactor = DataStorage.GetSpeedCorrectionFactor(vehicleGroup.AverageSpeed);
            
            emission += specificEmission * vehicleGroup.MaxTrafficIntensity * speedCorrectionFactor;
        }
        
        emission *= length / 3600f;

        var result = new EmissionsResult
        {
            MaximumEmission = emission,
            PollutantInfo =  pollutantInfo
        };
        
        return result;
    }
    
    private EmissionsResult? CalculateTrafficLightQueueEmissions(Pollutant pollutant, List<VehicleGroupQueue> vehicleGroups, int trafficLightCycles, float trafficLightStopTime)
    {
        var pollutantInfo = DataStorage.PollutantInfos.First(i => i.Pollutant == pollutant);
        
        if (!DataStorage.VehicleEmissionFactors[vehicleGroups.First().VehicleType].ContainsKey(pollutant))
        {
            return null;
        }
        
        var emission = 0f;

        foreach (var vehicleGroup in vehicleGroups)
        {
            var specificEmission = DataStorage.VehicleSpecificEmissions[vehicleGroup.VehicleType][pollutant];

            emission += specificEmission * vehicleGroup.VehiclesCount;
        }
        
        emission *= trafficLightCycles * trafficLightStopTime / 40f;

        var result = new EmissionsResult
        {
            MaximumEmission = emission,
            PollutantInfo =  pollutantInfo
        };
        
        return result;
    }

    #endregion
}