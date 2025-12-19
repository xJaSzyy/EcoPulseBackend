using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models.VehicleFlowEmissionSource;

public class VehicleFlowEmissionSourceAddModel
{
    /// <summary>
    /// Начальные координаты
    /// </summary>
    public Coordinates StartLocation { get; set; } = null!;
    
    /// <summary>
    /// Конечные координаты
    /// </summary>
    public Coordinates EndLocation { get; set; } = null!;
    
    /// <summary>
    /// Тип транспортного средства
    /// </summary>
    public VehicleType VehicleType { get; set; }
        
    /// <summary>
    /// Фактическая наибольшая интенсивность движения
    /// </summary>
    public float MaxTrafficIntensity { get; set; }
        
    /// <summary>
    /// Средняя скорость движения транспортного потока
    /// </summary>
    public float AverageSpeed { get; set; }
}