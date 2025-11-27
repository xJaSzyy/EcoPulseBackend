namespace EcoPulseBackend.Models.Calculate;

public class VehicleFlowEmissionsCalculateModel
{
    /// <summary>
    /// Список групп транспортных средств
    /// </summary>
    public List<VehicleGroup> VehicleGroups { get; set; } = new();
    
    /// <summary>
    /// Протяженность автомагистрали (или ее участка)
    /// </summary>
    public float Length { get; set; }
}