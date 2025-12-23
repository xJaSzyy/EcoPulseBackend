using System.ComponentModel.DataAnnotations;
using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models.VehicleFlowEmissionSource;

/// <summary>
/// Источник выбросов из потока транспортных средств
/// </summary>
public class VehicleFlowEmissionSource
{
    /// <summary>
    /// Идентификатор
    /// </summary>
    [Key]
    public int Id { get; set; }
    
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