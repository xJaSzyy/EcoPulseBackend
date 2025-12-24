using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DangerZone;
using EcoPulseBackend.Models.MaximumSingle;

namespace EcoPulseBackend.Services;

public class MaximumSingleService : IMaximumSingleService
{
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
    
    public EmissionsGroupResult CalculateEmissions(MaximumSingleEmissionsCalculateModel model)
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
    
    public SingleDangerZone CalculateDangerZone(MaximumSingleEmissionsCalculateModel model)
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
}