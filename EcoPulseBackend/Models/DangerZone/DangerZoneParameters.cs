namespace EcoPulseBackend.Models.DangerZone;

/// <summary>
/// Параметры зоны выброса
/// </summary>
public class DangerZoneParameters
{
    /// <summary>
    /// Идентификатор источника выброса
    /// </summary>
    public int EmissionSourceId { get; set; }
    
    /// <summary>
    /// Долгота
    /// </summary>
    public double Lon { get; set; }
    
    /// <summary>
    /// Широта
    /// </summary>
    public double Lat { get; set; }
    
    /// <summary>
    /// Длина зоны выброса
    /// </summary>
    public double Length { get; set; }
    
    /// <summary>
    /// Ширина зоны выброса
    /// </summary>
    public double Width { get; set; }
    
    /// <summary>
    /// Цвет зоны выброса
    /// </summary>
    public string Color { get; set; } = null!;
    
    /// <summary>
    /// Среднее значение из n макисмальных концентраций
    /// </summary>
    public float AverageConcentration { get; set; }
    
    /// <summary>
    /// Угол направления
    /// </summary>
    public double Angle { get; set; }
}