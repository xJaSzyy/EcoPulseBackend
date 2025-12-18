using System.ComponentModel.DataAnnotations;

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
    /// Начальные координаты
    /// </summary>
    public Coordinates StartLocation { get; set; } = null!;
    
    /// <summary>
    /// Конечные координаты
    /// </summary>
    public Coordinates EndLocation { get; set; } = null!;
}