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
    /// Долгота
    /// </summary>
    public float Lon { get; set; }
    
    /// <summary>
    /// Широта
    /// </summary>
    public float Lat { get; set; }
}