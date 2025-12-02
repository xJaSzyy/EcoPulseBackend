using EcoPulseBackend.Enums;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.Calculate;
using EcoPulseBackend.Services;

namespace Tests;

public class Tests
{
    private IEmissionService _service;
    
    [SetUp]
    public void Setup()
    {
        _service = new EmissionService();
    }

    [Test]
    public void CalculateGasolineGeneratorEmissionsBatch_ReturnsValidResults()
    {
        // Arrange
        var model = new GasolineGeneratorEmissionsCalculateModel
        {
            WorkHoursPerDay = 1,
            WorkDaysPerYear = 365,
            GeneratorCount = 1,
            SameGeneratorCount = 1
        };

        var expectedResult = new List<EmissionsResult>
        {
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.NO2 },
                MaximumEmission = 0.000039f, 
                GrossEmission = 0.000051f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.NO },
                MaximumEmission = 0.000006f, 
                GrossEmission = 0.000008f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.SO2 },
                MaximumEmission = 0.000012f, 
                GrossEmission = 0.000016f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.CO },
                MaximumEmission = 0.002604f,
                GrossEmission = 0.003422f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.CH },
                MaximumEmission = 0.000347f, 
                GrossEmission = 0.000456f
            },
        };

        // Act
        var actualResult = _service.CalculateGasolineGeneratorEmissionsBatch(model);

        // Assert
        Assert.That(actualResult, Has.Count.EqualTo(5));

        foreach (var actual in actualResult)
        {
            var expected = expectedResult.FirstOrDefault(x => x.PollutantInfo.Pollutant == actual.PollutantInfo.Pollutant);
            
            Assert.That(expected, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That((float)Math.Round(actual.MaximumEmission, 6), Is.EqualTo(expected.MaximumEmission));
                Assert.That((float)Math.Round(actual.GrossEmission, 6), Is.EqualTo(expected.GrossEmission));
            });
        }
    }
}