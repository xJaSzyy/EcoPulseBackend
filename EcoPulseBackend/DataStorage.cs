using EcoPulseBackend.Enums;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.Reservoirs;

namespace EcoPulseBackend;

public static class DataStorage
{
    public static readonly List<PollutantInfo> PollutantInfos = new()
    {
        new PollutantInfo
        {
            Code = 337,
            Name = "Углерода оксид (углерод окись; углерод моноокись; угарный газ)",
            Pollutant = Pollutant.CO,
            SpecificEmission = 7.5f,
            MaxPermissibleConcentration = 5f,
            DailyAverageConcentration = 3f
        },
        new PollutantInfo
        {
            Code = 2704,
            Name = "Бензин (нефтяной, малосернистый) /в пересчете на углерод/",
            Pollutant = Pollutant.CH,
            SpecificEmission = 1.0f,
            MaxPermissibleConcentration = 5f,
            DailyAverageConcentration = 1.5f
        },
        new PollutantInfo
        {
            Code = 301,
            Name = "Азота диоксид (двуокись азота; пероксид азота)",
            Pollutant = Pollutant.NO2,
            SpecificEmission = 0.112f,
            Mass = 0.2695f,
            MaxPermissibleConcentration = 0.2f,
            DailyAverageConcentration = 0.04f
        },
        new PollutantInfo
        {
            Code = 304,
            Name = "Азота оксид (азот (II) оксид; азот монооксид)",
            Pollutant = Pollutant.NO,
            SpecificEmission = 0.0182f,
            Mass = 0.0444f,
            MaxPermissibleConcentration = 0.4f,
            DailyAverageConcentration = 0.06f
        },
        new PollutantInfo
        {
            Code = 330,
            Name = "Серы диоксид",
            Pollutant = Pollutant.SO2,
            SpecificEmission = 0.036f,
            Mass = 1.0528f,
            MaxPermissibleConcentration = 0.5f,
            DailyAverageConcentration = 0.05f
        },
        new PollutantInfo
        {
            Code = 2754,
            Name = "Углеводороды предельные C12 - C19 (растворители РПК-240, РПК-280)",
            ShortName = "углеводороды",
            Pollutant = Pollutant.RPK240280,
            SpecificEmission = 99.72f,
            MaxPermissibleConcentration = 1f
        },
        new PollutantInfo
        {
            Code = 333,
            Name = "Сероводород (дигидросульфид; водород сернистый; гидросульфид)",
            ShortName = "дигидросульфид",
            Pollutant = Pollutant.H2S,
            SpecificEmission = 0.28f,
            MaxPermissibleConcentration = 0.008f
        },
        new PollutantInfo
        {
            Code = 123,
            Name = "диЖелезо триоксид (железа оксид; железо сесквиоксид) /в пересчете на железо/",
            Pollutant = Pollutant.Fe2O3,
            DailyAverageConcentration = 0.04f
        },
        new PollutantInfo
        {
            Code = 143,
            Name = "Марганец и его соединения /в пересчете на марганец (IV) оксид/",
            Pollutant = Pollutant.MnO2,
            MaxPermissibleConcentration = 0.01f,
            DailyAverageConcentration = 0.001f
        },
        new PollutantInfo
        {
            Code = 342,
            Name = "Фториды газообразные /в пересчете на фтор/: гидрофторид (водород фторид, фторводород); кремний тетрафторид",
            Pollutant = Pollutant.FluorideGases,
            MaxPermissibleConcentration = 0.02f,
            DailyAverageConcentration = 0.005f
        },
        new PollutantInfo
        {
            Code = 380,
            Name = "Углерод диоксид",
            Pollutant = Pollutant.CO2,
            Mass = 4.9f,
            MaxPermissibleConcentration = 5f
        },
        new PollutantInfo
        {
            Code = 2,
            Name = "Твердые частицы",
            Pollutant = Pollutant.SP,
            Mass = 15.72f,
            MaxPermissibleConcentration = 0.5f
        },
        new PollutantInfo
        {
            Code = 328,
            Name = "Сажа",
            Pollutant = Pollutant.Soot
        },
        new PollutantInfo
        {
            Code = 184,
            Name = "Соединения свинца",
            Pollutant = Pollutant.LeadCompounds,
            MaxPermissibleConcentration = 0.001f,
            DailyAverageConcentration = 0.0003f
        },
        new PollutantInfo
        {
            Code = 1325,
            Name = "Формальдегид",
            Pollutant = Pollutant.CH2O,
            MaxPermissibleConcentration = 0.035f,
            DailyAverageConcentration = 0.003f
        },
        new PollutantInfo
        {
            Code = 703,
            Name = "Бенз(а)пирен",
            Pollutant = Pollutant.C20H12,
        },
        new PollutantInfo
        {
            Code = 3749,
            Name = "Пыль каменного угля",
            Pollutant = Pollutant.CoalDust,
        },
    };

    public static readonly Dictionary<ReservoirType, Dictionary<ClimateZone, Dictionary<OilProduct, VaporConcentrationRecord>>>
        VaporConcentration = new()
        {
            {
                ReservoirType.Ground, new Dictionary<ClimateZone, Dictionary<OilProduct, VaporConcentrationRecord>>
                {
                    {
                        ClimateZone.First, new Dictionary<OilProduct, VaporConcentrationRecord>
                        {
                            { OilProduct.AutomobileGasoline, new VaporConcentrationRecord(464f, 205f, 248f) },
                            { OilProduct.DieselFuel, new VaporConcentrationRecord(1.49f, 0.79f, 1.06f) },
                            { OilProduct.Oils, new VaporConcentrationRecord(0.16f, 0.10f, 0.10f) }
                        }
                    },
                    {
                        ClimateZone.Second, new Dictionary<OilProduct, VaporConcentrationRecord>
                        {
                            { OilProduct.AutomobileGasoline, new VaporConcentrationRecord(580f, 250f, 310f) },
                            { OilProduct.DieselFuel, new VaporConcentrationRecord(1.86f, 0.96f, 1.32f) },
                            { OilProduct.Oils, new VaporConcentrationRecord(0.20f, 0.12f, 0.12f) }
                        }
                    },
                    {
                        ClimateZone.Third, new Dictionary<OilProduct, VaporConcentrationRecord>
                        {
                            { OilProduct.AutomobileGasoline, new VaporConcentrationRecord(701.8f, 310f, 375.1f) },
                            { OilProduct.DieselFuel, new VaporConcentrationRecord(2.25f, 1.19f, 1.60f) },
                            { OilProduct.Oils, new VaporConcentrationRecord(0.24f, 0.15f, 0.15f) }
                        }
                    }
                }
            },
            {
                ReservoirType.Buried, new Dictionary<ClimateZone, Dictionary<OilProduct, VaporConcentrationRecord>>
                {
                    {
                        ClimateZone.First, new Dictionary<OilProduct, VaporConcentrationRecord>
                        {
                            { OilProduct.AutomobileGasoline, new VaporConcentrationRecord(384f, 172.2f, 255f) },
                            { OilProduct.DieselFuel, new VaporConcentrationRecord(1.24f, 0.66f, 0.88f) },
                            { OilProduct.Oils, new VaporConcentrationRecord(0.13f, 0.08f, 0.08f) }
                        }
                    },
                    {
                        ClimateZone.Second, new Dictionary<OilProduct, VaporConcentrationRecord>
                        {
                            { OilProduct.AutomobileGasoline, new VaporConcentrationRecord(480f, 210.2f, 255f) },
                            { OilProduct.DieselFuel, new VaporConcentrationRecord(1.55f, 0.80f, 1.10f) },
                            { OilProduct.Oils, new VaporConcentrationRecord(0.16f, 0.10f, 0.10f) }
                        }
                    },
                    {
                        ClimateZone.Third, new Dictionary<OilProduct, VaporConcentrationRecord>
                        {
                            { OilProduct.AutomobileGasoline, new VaporConcentrationRecord(508f, 260.4f, 308.5f) },
                            { OilProduct.DieselFuel, new VaporConcentrationRecord(1.88f, 0.99f, 1.33f) },
                            { OilProduct.Oils, new VaporConcentrationRecord(0.19f, 0.12f, 0.12f) }
                        }
                    }
                }
            }
        };
    
    public static readonly Dictionary<MetalMachiningMachineType, float> SpecificDustEmissionsByType = new()
    {
        { MetalMachiningMachineType.Drilling, 0.007f },
        { MetalMachiningMachineType.Milling, 0.097f },
        { MetalMachiningMachineType.Cutting, 0.203f }
    };
    
    public static readonly Dictionary<Pollutant, float> SpecificEmissionsByElectrodes = new()
    {
        { Pollutant.Fe2O3, 9.77f },
        { Pollutant.MnO2, 1.73f },
        { Pollutant.FluorideGases, 0.40f }
    };

    public static readonly Dictionary<VehicleType, Dictionary<Pollutant, float>> VehicleEmissionFactors = new()
    {
        {
            VehicleType.Passenger, new Dictionary<Pollutant, float>
            {
                { Pollutant.CO, 19f },
                { Pollutant.NO2, 1.8f },
                { Pollutant.CH, 2.1f },
                { Pollutant.SO2, 0.065f },
                { Pollutant.CH2O, 0.006f },
                { Pollutant.LeadCompounds, 0.019f },
                { Pollutant.C20H12, 1.7f * 1e-6f },
            }
        },
        {
            VehicleType.DieselPassenger, new Dictionary<Pollutant, float>
            {
                { Pollutant.CO, 2f },
                { Pollutant.NO2, 1.3f },
                { Pollutant.CH, 0.25f },
                { Pollutant.Soot, 0.1f },
                { Pollutant.SO2, 0.21f },
                { Pollutant.CH2O, 0.003f },
            }
        },
        {
            VehicleType.CargoCarburetorLow, new Dictionary<Pollutant, float>()
            {
                { Pollutant.CO, 69.4f },
                { Pollutant.NO2, 2.9f },
                { Pollutant.CH, 11.5f },
                { Pollutant.SO2, 0.2f },
                { Pollutant.CH2O, 0.02f },
                { Pollutant.LeadCompounds, 0.026f },
                { Pollutant.C20H12, 4.5f * 1e-6f },
            }
        },
        {
            VehicleType.CargoCarburetorHigh, new Dictionary<Pollutant, float>()
            {
                { Pollutant.CO, 75f },
                { Pollutant.NO2, 5.2f },
                { Pollutant.CH, 13.4f },
                { Pollutant.SO2, 0.22f },
                { Pollutant.CH2O, 0.022f },
                { Pollutant.LeadCompounds, 0.033f },
                { Pollutant.C20H12, 6.3f * 1e-6f },
            }
        },
        {
            VehicleType.CarburetorBuses, new Dictionary<Pollutant, float>()
            {
                { Pollutant.CO, 97.6f },
                { Pollutant.NO2, 5.3f },
                { Pollutant.CH, 13.4f },
                { Pollutant.SO2, 0.32f },
                { Pollutant.CH2O, 0.03f },
                { Pollutant.LeadCompounds, 0.041f },
                { Pollutant.C20H12, 6.4f * 1e-6f },
            }
        },
        {
            VehicleType.DieselTrucks, new Dictionary<Pollutant, float>()
            {
                { Pollutant.CO, 8.5f },
                { Pollutant.NO2, 7.7f },
                { Pollutant.CH, 6f },
                { Pollutant.Soot, 0.3f },
                { Pollutant.SO2, 1.25f },
                { Pollutant.CH2O, 0.21f },
                { Pollutant.C20H12, 6.5f * 1e-6f },
            }
        },
        {
            VehicleType.DieselBuses, new Dictionary<Pollutant, float>()
            {
                { Pollutant.CO, 8.8f },
                { Pollutant.NO2, 8f },
                { Pollutant.CH, 6.5f },
                { Pollutant.Soot, 0.3f },
                { Pollutant.SO2, 1.45f },
                { Pollutant.CH2O, 0.31f },
                { Pollutant.C20H12, 6.7f * 1e-6f },
            }
        },
        {
            VehicleType.CargoGas, new Dictionary<Pollutant, float>()
            {
                { Pollutant.CO, 39f },
                { Pollutant.NO2, 2.6f },
                { Pollutant.CH, 1.3f },
                { Pollutant.SO2, 0.18f },
                { Pollutant.CH2O, 0.002f },
                { Pollutant.C20H12, 2f * 1e-6f },
            }
        }
    };

    private static readonly Dictionary<int, float> SpeedCorrectionFactors = new()
    {
        { 10, 1.35f },
        { 15, 1.28f },
        { 20, 1.2f },
        { 25, 1.1f },
        { 30, 1f },
        { 35, 0.88f },
        { 40, 0.75f },
        { 45, 0.63f },
        { 50, 0.5f },
        { 60, 0.3f },
        { 75, 0.45f },
        { 80, 0.5f },
        { 100, 0.65f }
    };
    
    public static readonly Dictionary<VehicleType, Dictionary<Pollutant, float>> VehicleSpecificEmissions = new()
    {
        {
            VehicleType.Passenger, new Dictionary<Pollutant, float>
            {
                { Pollutant.CO, 3.5f },
                { Pollutant.NO2, 0.05f },
                { Pollutant.CH, 0.25f },
                { Pollutant.SO2, 0.01f },
                { Pollutant.CH2O, 0.0008f },
                { Pollutant.LeadCompounds, 0.0044f },
                { Pollutant.C20H12, 2f * 1e-6f },
            }
        },
        {
            VehicleType.DieselPassenger, new Dictionary<Pollutant, float>
            {
                { Pollutant.CO, 0.13f },
                { Pollutant.NO2, 0.08f },
                { Pollutant.CH, 0.06f },
                { Pollutant.Soot, 0.035f },
                { Pollutant.SO2, 0.04f },
                { Pollutant.CH2O, 0.0008f },
            }
        },
        {
            VehicleType.CargoCarburetorLow, new Dictionary<Pollutant, float>()
            {
                { Pollutant.CO, 6.3f },
                { Pollutant.NO2, 0.075f },
                { Pollutant.CH, 1f },
                { Pollutant.SO2, 0.02f },
                { Pollutant.CH2O, 0.0015f },
                { Pollutant.LeadCompounds, 0.0047f },
                { Pollutant.C20H12, 4f * 1e-6f },
            }
        },
        {
            VehicleType.CargoCarburetorHigh, new Dictionary<Pollutant, float>()
            {
                { Pollutant.CO, 18.4f },
                { Pollutant.NO2, 0.2f },
                { Pollutant.CH, 2.96f },
                { Pollutant.SO2, 0.028f },
                { Pollutant.CH2O, 0.006f },
                { Pollutant.LeadCompounds, 0.0075f },
                { Pollutant.C20H12, 4.5f * 1e-6f },
            }
        },
        {
            VehicleType.CarburetorBuses, new Dictionary<Pollutant, float>()
            {
                { Pollutant.CO, 16.1f },
                { Pollutant.NO2, 0.16f },
                { Pollutant.CH, 2.64f },
                { Pollutant.SO2, 0.03f },
                { Pollutant.CH2O, 0.012f },
                { Pollutant.LeadCompounds, 0.0075f },
                { Pollutant.C20H12, 4.5f * 1e-6f },
            }
        },
        {
            VehicleType.DieselTrucks, new Dictionary<Pollutant, float>()
            {
                { Pollutant.CO, 2.85f },
                { Pollutant.NO2, 0.81f },
                { Pollutant.CH, 0.3f },
                { Pollutant.Soot, 0.07f },
                { Pollutant.SO2, 0.075f },
                { Pollutant.CH2O, 0.015f },
                { Pollutant.C20H12, 6.3f * 1e-6f },
            }
        },
        {
            VehicleType.DieselBuses, new Dictionary<Pollutant, float>()
            {
                { Pollutant.CO, 3.07f },
                { Pollutant.NO2, 0.7f },
                { Pollutant.CH, 0.41f },
                { Pollutant.Soot, 0.09f },
                { Pollutant.SO2, 0.09f },
                { Pollutant.CH2O, 0.02f },
                { Pollutant.C20H12, 6.4f * 1e-6f },
            }
        },
        {
            VehicleType.CargoGas, new Dictionary<Pollutant, float>()
            {
                { Pollutant.CO, 6.44f },
                { Pollutant.NO2, 0.09f },
                { Pollutant.CH, 0.26f },
                { Pollutant.SO2, 0.01f },
                { Pollutant.CH2O, 0.0004f },
                { Pollutant.C20H12, 3.6f * 1e-6f },
            }
        }
    };
    
    public static float GetSpeedCorrectionFactor(double speed)
    {
        var nearest = SpeedCorrectionFactors
            .Where(x => x.Key >= speed)
            .OrderBy(x => x.Key)
            .FirstOrDefault();
    
        return nearest.Key == 0 ? SpeedCorrectionFactors.Values.Last() : nearest.Value;
    }
    
    public static readonly SortedDictionary<double, string> ColorMap = new()
    {
        { 225.4, "rgba(164, 125, 184, 1)" },
        { 125.4, "rgba(246, 104, 106, 1)" },
        { 55.4, "rgba(251, 153, 86, 1)" },
        { 35.4, "rgba(248, 212, 97, 1)" },
        { 9.0, "rgba(171, 209, 98, 1)" }
    };
}