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
    public float? SpecificEmission { get; set; }
    
    /// <summary>
    /// Масса
    /// </summary>
    public float? Mass { get; set; }
    
    /// <summary>
    /// Предельно допустимая максимальная разовая концентрация
    /// </summary>
    public float? MaxPermissibleConcentration { get; set; }
    
    /// <summary>
    /// Предельно допустимая среднесуточная концентрация
    /// </summary>
    public float? DailyAverageConcentration { get; set; }
}