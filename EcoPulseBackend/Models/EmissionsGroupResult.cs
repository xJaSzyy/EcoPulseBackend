namespace EcoPulseBackend.Models;

public class EmissionsGroupResult
{
    public PollutantInfo PollutantInfo { get; set; } = null!;
    
    public List<EmissionsResult> Emissions { get; set; } = new();
}