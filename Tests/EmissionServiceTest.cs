using EcoPulseBackend.Enums;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DuringMetalMachining;
using EcoPulseBackend.Models.DuringWeldingOperations;
using EcoPulseBackend.Models.GasolineGenerator;
using EcoPulseBackend.Models.MaximumSingle;
using EcoPulseBackend.Models.OpenCoalWarehouse;
using EcoPulseBackend.Models.Reservoirs;
using EcoPulseBackend.Models.TrafficLightQueue;
using EcoPulseBackend.Models.VehicleFlow;
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
    public void CalculateGasolineGeneratorEmissionsBatch_ReturnsValidResult()
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
        Assert.That(actualResult, Has.Count.EqualTo(expectedResult.Count));

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

    [Test]
    public void CalculateReservoirsEmissionsBatch_ReturnsValidResult()
    {
        // Arrange
        var model = new ReservoirsEmissionsCalculateModel
        {
            ReservoirType = ReservoirType.Ground,
            OilProduct = OilProduct.DieselFuel,
            ClimateZone = ClimateZone.Second,
            AutumnWinterOilAmount = 100f,
            SpringSummerOilAmount = 50f,
            DrainedVolume = 150f,
            AverageDrainTime = 1200f
        };

        var expectedResult = new ReservoirsEmissionsBatchResult
        {
            AnnualInjectionEmissions = 0.000162f,
            AnnualIrrigationEmissions = 0.0075f,
            Emissions = new List<EmissionsResult>
            {
                new()
                {
                    PollutantInfo = new PollutantInfo { Pollutant = Pollutant.RPK240280 },
                    MaximumEmission = 0.231849f, 
                    GrossEmission = 0.007641f
                },
                new()
                {
                    PollutantInfo = new PollutantInfo { Pollutant = Pollutant.H2S },
                    MaximumEmission = 0.000651f, 
                    GrossEmission = 0.000021f
                },
            }
        };

        // Act
        var actualResult = _service.CalculateReservoirsEmissionsBatch(model);
        
        // Assert
        Assert.Multiple(() =>
        {

            Assert.That(actualResult.Emissions, Has.Count.EqualTo(expectedResult.Emissions.Count));
            Assert.That(actualResult.AnnualInjectionEmissions, Is.EqualTo(expectedResult.AnnualInjectionEmissions));
            Assert.That(actualResult.AnnualIrrigationEmissions, Is.EqualTo(expectedResult.AnnualIrrigationEmissions));
        });
        
        foreach (var actual in actualResult.Emissions)
        {
            var expected = expectedResult.Emissions.FirstOrDefault(x => x.PollutantInfo.Pollutant == actual.PollutantInfo.Pollutant);
            
            Assert.That(expected, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That((float)Math.Round(actual.MaximumEmission, 6), Is.EqualTo(expected.MaximumEmission));
                Assert.That((float)Math.Round(actual.GrossEmission, 6), Is.EqualTo(expected.GrossEmission));
            });
        }
    }
    
    [TestCase(MetalMachiningMachineType.Drilling, 0.0014f, 0.001840f)]
    [TestCase(MetalMachiningMachineType.Milling,0.0194f, 0.025492f)]
    [TestCase(MetalMachiningMachineType.Cutting,0.0406f, 0.053348f)]
    public void CalculateDuringMetalMachiningEmissionsBatch_ReturnsValidResult(MetalMachiningMachineType type, float maximumEmission, float grossEmission)
    {
        // Arrange
        var model = new DuringMetalMachiningEmissionsCalculateModel
        {
            MetalMachiningMachineType = type,
            WorkDaysPerYear = 365
        };

        var expectedResult = new List<EmissionsResult>
        {
            new EmissionsResult
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.Fe2O3 },
                MaximumEmission = maximumEmission,
                GrossEmission = grossEmission
            }
        };

        // Act
        var actualResult = _service.CalculateDuringMetalMachiningEmissionsBatch(model);
        
        // Assert
        Assert.That(actualResult, Has.Count.EqualTo(expectedResult.Count));
        
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
    
    [Test]
    public void CalculateDuringWeldingOperationsEmissionsBatch_ReturnsValidResult()
    {
        // Arrange
        var model = new DuringWeldingOperationsEmissionsCalculateModel
        {
            ElectrodesPerYear = 241.36f, 
            WorkDaysPerYear = 365
        };

        var expectedResult = new DuringWeldingOperationsEmissionsBatchResult
        {
            NormElectrodesPerYear = 205.16f,
            MaterialsConsumption = 0.56f,
            Emissions = new List<EmissionsResult>
            {
                new()
                {
                    PollutantInfo = new PollutantInfo { Pollutant = Pollutant.Fe2O3 },
                    MaximumEmission = 0.00061f, 
                    GrossEmission = 0.000802f
                },
                new()
                {
                    PollutantInfo = new PollutantInfo { Pollutant = Pollutant.MnO2 },
                    MaximumEmission = 0.000108f, 
                    GrossEmission = 0.000142f
                },
                new()
                {
                    PollutantInfo = new PollutantInfo { Pollutant = Pollutant.FluorideGases },
                    MaximumEmission = 0.000025f, 
                    GrossEmission = 0.000033f
                }
            }
        };

        // Act
        var actualResult = _service.CalculateDuringWeldingOperationsEmissionsBatch(model);
        
        // Assert
        Assert.Multiple(() =>
        {

            Assert.That(actualResult.Emissions, Has.Count.EqualTo(expectedResult.Emissions.Count));
            Assert.That((float)Math.Round(actualResult.NormElectrodesPerYear, 2), Is.EqualTo(expectedResult.NormElectrodesPerYear));
            Assert.That((float)Math.Round(actualResult.MaterialsConsumption, 2), Is.EqualTo(expectedResult.MaterialsConsumption));
        });
        
        foreach (var actual in actualResult.Emissions)
        {
            var expected = expectedResult.Emissions.FirstOrDefault(x => x.PollutantInfo.Pollutant == actual.PollutantInfo.Pollutant);
            
            Assert.That(expected, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That((float)Math.Round(actual.MaximumEmission, 6), Is.EqualTo(expected.MaximumEmission));
                Assert.That((float)Math.Round(actual.GrossEmission, 6), Is.EqualTo(expected.GrossEmission));
            });
        }
    }

    [Test]
    public void CalculateMaximumSingleEmissions_ReturnsValidResult()
    {
        // Arrange
        var model = new MaximumSingleEmissionsCalculateModel
        {
            Pollutant = Pollutant.SP,
            EjectedTemp = 235,
            AirTemp = -40,
            AvgExitSpeed = 15,
            HeightSource = 13,
            DiameterSource = 1,
            TempStratificationRatio = CoefficientRegion.CentralRegions,
            SedimentationRateRatio = CoefficientDegreePurification.Medium,
            Distance = 15,
            MaxCount = 3
        };

        var expectedResult = new EmissionsGroupResult
        {
            PollutantInfo = new PollutantInfo { Pollutant = Pollutant.SP },
            Emissions = new List<EmissionsResult>
            {
                new()
                {
                    MaximumEmission = 0.006019f
                },
                new()
                {
                    MaximumEmission = 0.023284f
                },
                new()
                {
                    MaximumEmission = 0.050637f
                }
            }
        };

        // Act
        var actualResult = _service.CalculateMaximumSingleEmissions(model);

        // Assert
        Assert.That(actualResult.Emissions, Has.Count.EqualTo(expectedResult.Emissions.Count));

        foreach (var actual in actualResult.Emissions)
        {
            var expected = expectedResult.Emissions[actualResult.Emissions.IndexOf(actual)];
            
            Assert.That(expected, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That((float)Math.Round(actual.MaximumEmission, 6), Is.EqualTo(expected.MaximumEmission));
            });
        }
    }
    
    [Test]
    public void CalculateVehicleFlowEmissionsBatch_ReturnsValidResult()
    {
        // Arrange
        var model = new VehicleFlowEmissionsCalculateModel
        {
            VehicleGroups = new List<VehicleGroup>
            {
                new()
                {
                    VehicleType = VehicleType.Passenger,
                    AverageSpeed = 90,
                    MaxTrafficIntensity = 3
                }
            },
            Length = 123
        };
        
        var expectedResult = new List<EmissionsResult>
        {
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.LeadCompounds },
                MaximumEmission = 0.001265875f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.NO2 },
                MaximumEmission = 0.119924985f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.SO2 },
                MaximumEmission = 0.004330625f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.CO },
                MaximumEmission = 1.265875f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.C20H12 },
                MaximumEmission = 1.132625e-7f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.CH2O },
                MaximumEmission = 0.00039974996f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.CH },
                MaximumEmission = 0.1399125f
            }
        };

        // Act
        var actualResult = _service.CalculateVehicleFlowEmissionsBatch(model);

        // Assert
        Assert.That(actualResult, Has.Count.EqualTo(expectedResult.Count));

        foreach (var actual in actualResult)
        {
            var expected = expectedResult.FirstOrDefault(x => x.PollutantInfo.Pollutant == actual.PollutantInfo.Pollutant);
            
            Assert.That(expected, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(actual.MaximumEmission, Is.EqualTo(expected.MaximumEmission));
                Assert.That(actual.GrossEmission, Is.EqualTo(expected.GrossEmission));
            });
        }
    }
    
    [Test]
    public void CalculateTrafficLightQueueEmissionsBatch_ReturnsValidResult()
    {
        // Arrange
        var model = new TrafficLightQueueEmissionsCalculateModel
        {
            VehicleGroups = new List<VehicleGroupQueue>
            {
                new()
                {
                    VehicleType = VehicleType.Passenger,
                    VehiclesCount = 12
                }
            },
            TrafficLightCycles = 20,
            TrafficLightStopTime = 60f
        };
        
        var expectedResult = new List<EmissionsResult>
        {
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.LeadCompounds },
                MaximumEmission = 1.58399999f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.NO2 },
                MaximumEmission = 18.0f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.SO2 },
                MaximumEmission = 3.5999999f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.CO },
                MaximumEmission = 1260.0f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.C20H12 },
                MaximumEmission = 0.000720000011f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.CH2O },
                MaximumEmission = 0.287999988f
            },
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.CH },
                MaximumEmission = 90.0f
            }
        };

        // Act
        var actualResult = _service.CalculateTrafficLightQueueEmissionsBatch(model);

        // Assert
        Assert.That(actualResult, Has.Count.EqualTo(expectedResult.Count));

        foreach (var actual in actualResult)
        {
            var expected = expectedResult.FirstOrDefault(x => x.PollutantInfo.Pollutant == actual.PollutantInfo.Pollutant);
            
            Assert.That(expected, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(actual.MaximumEmission, Is.EqualTo(expected.MaximumEmission));
                Assert.That(actual.GrossEmission, Is.EqualTo(expected.GrossEmission));
            });
        }
    }
    
    [Test]
    public void CalculateOpenCoalWarehouseEmissions_ReturnsValidResult()
    {
        // Arrange
        var model = new OpenCoalWarehouseEmissionsCalculateModel
        {
            SpecificEmission = 0.32f,
            UnloadMaterialCountPerYear = 2700000f,
            UnloadMaterialCountPerHour = 285.388f,
            DustSuppressionEfficiency = 0f,
            CoalPileBaseArea = 1700f,
            SnowyDaysCount = 162,
            RainyDaysCount = 89
        };

        var expectedResult = new List<EmissionsResult>
        {
            new()
            {
                PollutantInfo = new PollutantInfo { Pollutant = Pollutant.CoalDust },
                MaximumEmission = 0.368804f, 
                GrossEmission = 3.853858f
            }
        };

        // Act
        var actualResult = _service.CalculateOpenCoalWarehouseEmissions(model);

        // Assert
        Assert.That(actualResult, Has.Count.EqualTo(expectedResult.Count));

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