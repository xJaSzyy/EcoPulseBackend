using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models.VehicleFlowEmissionSource;

public class VehicleFlowEmissionSourceAddModel
{
    /// <summary>
    /// Идентификатор города
    /// </summary>
    public int CityId { get; set; }
    
    /// <summary>
    /// Список координат
    /// </summary>
    public List<Coordinates> Points { get; set; } = null!;
    
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