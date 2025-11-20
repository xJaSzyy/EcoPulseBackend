namespace EcoPulseBackend.Models;

/// <summary>
/// Концентрация паров нефтепродуктов в выбросах паровоздушной смеси при заполнении резервуаров и баков автомашин
/// </summary>
public class VaporConcentrationRecord
{
    public VaporConcentrationRecord(float maxVaporConcentration, float autumnWinterVaporConcentration, float springSummerVaporConcentration)
    {
        MaxVaporConcentration = maxVaporConcentration;
        AutumnWinterVaporConcentration = autumnWinterVaporConcentration;
        SpringSummerVaporConcentration = springSummerVaporConcentration;
    }

    public float MaxVaporConcentration { get; set; }
    
    public float AutumnWinterVaporConcentration { get; set; }
    
    public float SpringSummerVaporConcentration { get; set; }
}