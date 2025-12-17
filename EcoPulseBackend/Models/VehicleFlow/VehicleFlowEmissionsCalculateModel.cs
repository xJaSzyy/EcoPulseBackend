namespace EcoPulseBackend.Models.VehicleFlow;

/// <summary>
/// Модель для расчета выбросов движущегося автотранспорта
/// </summary>
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