namespace EcoPulseBackend.Models;

/// <summary>
/// Параметры зоны выброса
/// </summary>
public class DangerZoneParameters
{
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
    public double AverageConcentration { get; set; }
}