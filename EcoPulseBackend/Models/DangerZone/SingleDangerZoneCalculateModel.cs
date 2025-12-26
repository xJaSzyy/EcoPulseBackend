using System.ComponentModel.DataAnnotations;
using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models.DangerZone;

public class SingleDangerZoneCalculateModel
{
    /// <summary>
    /// Загрязняющее вещество
    /// </summary>
    public Pollutant Pollutant { get; set; }
    
    /// <summary>
    /// Температура атмосферного воздуха
    /// </summary>
    [Range(-40, 40)]
    public float AirTemp { get; set; }
    
    /// <summary>
    /// Скорость ветра
    /// </summary>
    public float WindSpeed { get; set; }
    
    /// <summary>
    /// Направление ветра
    /// </summary>
    public float WindDirection { get; set; }
    
    /// <summary>
    /// Список идентификаторов городов
    /// </summary>
    public ICollection<int> CityIds { get; set; } = new List<int>();
}