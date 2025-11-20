using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models;

/// <summary>
/// Информация о ЗВ
/// </summary>
public class PollutantInfo
{
    /// <summary>
    /// Код
    /// </summary>
    public int Code { get; set; }

    /// <summary>
    /// Наименование
    /// </summary>
    public string Name { get; set; } = null!;
    
    /// <summary>
    /// Короткое наименование
    /// </summary>
    public string ShortName { get; set; } = null!;

    /// <summary>
    /// Загрязняющее вещество
    /// </summary>
    public Pollutant Pollutant { get; set; }
    
    /// <summary>
    /// Удельный выброс / концентрация
    /// </summary>
    public float SpecificEmission { get; set; }
}