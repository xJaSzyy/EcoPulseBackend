namespace EcoPulseBackend.Models;

public class DuringWeldingOperationsEmissionsBatchResult
{
    public float NormElectrodesPerYear { get; set; }
    
    public float MaterialsConsumption { get; set; }

    public List<EmissionsResult> Emissions { get; set; } = new();
}