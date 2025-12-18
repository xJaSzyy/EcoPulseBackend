using System.ComponentModel.DataAnnotations;
using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models.MaximumSingle;

/// <summary>
/// Модель для расчета выбросов загрязняющих вещество от одиночного точечного источника
/// </summary>
public class MaximumSingleEmissionsCalculateModel
{
    /// <summary>
    /// Загрязняющее вещество
    /// </summary>
    public Pollutant Pollutant { get; set; }
    
    /// <summary>
    /// Температура выбрасываемой ГВС
    /// </summary>
    [Range(235, 265)]
    public float EjectedTemp { get; set; }

    /// <summary>
    /// Температура атмосферного воздуха
    /// </summary>
    [Range(-40, 40)]
    public float AirTemp { get; set; }

    /// <summary>
    /// Средняя скорость выхода ГВС из устья источника выброса, м/с
    /// </summary>
    [Range(15, 30)]
    public float AvgExitSpeed { get; set; }

    /// <summary>
    /// Высота источника выброса, м.
    /// </summary>
    [Range(13, 150)]
    public float HeightSource { get; set; }

    /// <summary>
    /// Диаметр устья источника, м.
    /// </summary>
    [Range(1, 10)]
    public float DiameterSource { get; set; }

    /// <summary>
    /// Коэффициент региона
    /// </summary>
    public CoefficientRegion TempStratificationRatio { get; set; }

    /// <summary>
    /// Коэффициент степени очистки
    /// </summary>
    public CoefficientDegreePurification SedimentationRateRatio { get; set; }
    
    /// <summary>
    /// Скорость ветра
    /// </summary>
    public float WindSpeed { get; set; }

    /// <summary>
    /// Расстояние от источника выброса
    /// </summary>
    [Range(5, 10000)]
    public int Distance { get; set; }
    
    /// <summary>
    /// Количество максимальных точек
    /// </summary>
    public int MaxCount { get; set; }
}