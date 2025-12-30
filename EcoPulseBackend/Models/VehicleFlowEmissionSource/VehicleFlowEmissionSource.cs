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
    
    /// <summary>
    /// Идентификатор города
    /// </summary>
    public int CityId { get; set; }
    
    /// <summary>
    /// Город
    /// </summary>
    public City City { get; set; } = null!;
    
    /// <summary>
    /// Название улицы
    /// </summary>
    public string StreetName { get; set; } = null!;

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