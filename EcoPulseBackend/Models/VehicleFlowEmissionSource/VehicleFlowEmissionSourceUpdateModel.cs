using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models.VehicleFlowEmissionSource;

public class VehicleFlowEmissionSourceUpdateModel
{
    /// <summary>
    /// Идентификатор
    /// </summary>
    public int Id { get; set; }
    
    public List<Coordinates>? Points { get; set; }
    
    /// <summary>
    /// Тип транспортного средства
    /// </summary>
    public VehicleType? VehicleType { get; set; }
        
    /// <summary>
    /// Фактическая наибольшая интенсивность движения
    /// </summary>
    public float? MaxTrafficIntensity { get; set; }
        
    /// <summary>
    /// Средняя скорость движения транспортного потока
    /// </summary>
    public float? AverageSpeed { get; set; }
}