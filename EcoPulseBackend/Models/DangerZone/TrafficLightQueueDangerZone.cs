namespace EcoPulseBackend.Models.DangerZone;

public class TrafficLightQueueDangerZone
{
    /// <summary>
    /// Идентификатор источника выброса
    /// </summary>
    public int EmissionSourceId { get; set; }
    
    /// <summary>
    /// Координаты
    /// </summary>
    public Coordinates Location { get; set; } = null!;
    
    /// <summary>
    /// Цвет зоны выброса
    /// </summary>
    public string Color { get; set; } = null!;
    
    /// <summary>
    /// Среднее значение из n макисмальных концентраций
    /// </summary>
    public float AverageConcentration { get; set; }
}