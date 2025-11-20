using EcoPulseBackend.Enums;
using EcoPulseBackend.Models;

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
            SpecificEmission = 7.5f
        },
        new PollutantInfo
        {
            Code = 2704,
            Name = "Бензин (нефтяной, малосернистый) /в пересчете на углерод/",
            Pollutant = Pollutant.CH,
            SpecificEmission = 1.0f
        },
        new PollutantInfo
        {
            Code = 301,
            Name = "Азота диоксид (двуокись азота; пероксид азота)",
            Pollutant = Pollutant.NO2,
            SpecificEmission = 0.112f
        },
        new PollutantInfo
        {
            Code = 304,
            Name = "Азота оксид (азот (II) оксид; азот монооксид)",
            Pollutant = Pollutant.NO,
            SpecificEmission = 0.0182f
        },
        new PollutantInfo
        {
            Code = 330,
            Name = "Серы диоксид",
            Pollutant = Pollutant.SO2,
            SpecificEmission = 0.036f
        },
        new PollutantInfo
        {
            Code = 2754,
            Name = "Углеводороды предельные C12 - C19 (растворители РПК-240, РПК-280)",
            ShortName = "углеводороды",
            Pollutant = Pollutant.RPK240280,
            SpecificEmission = 99.72f
        },
        new PollutantInfo
        {
            Code = 333,
            Name = "Сероводород (дигидросульфид; водород сернистый; гидросульфид)",
            ShortName = "дигидросульфид",
            Pollutant = Pollutant.H2S,
            SpecificEmission = 0.28f
        },
        new PollutantInfo
        {
            Code = 123,
            Name = "диЖелезо триоксид (железа оксид; железо сесквиоксид) /в пересчете на железо/",
            Pollutant = Pollutant.Fe2O3
        },
        new PollutantInfo
        {
            Code = 143,
            Name = "Марганец и его соединения /в пересчете на марганец (IV) оксид/",
            Pollutant = Pollutant.MnO2
        },
        new PollutantInfo
        {
            Code = 342,
            Name = "Фториды газообразные /в пересчете на фтор/: гидрофторид (водород фторид, фторводород); кремний тетрафторид",
            Pollutant = Pollutant.FluorideGases
        }
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
}