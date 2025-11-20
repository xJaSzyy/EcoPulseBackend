namespace EcoPulseBackend.Models;

public class ReservoirsEmissionsBatchResult
{
    public float MaxVaporEmission { get; set; }
    
    public float AnnualInjectionEmissions { get; set; }
    
    public float AnnualIrrigationEmissions { get; set; }
    
    public List<EmissionsResult> Emissions { get; set; } = new();
}