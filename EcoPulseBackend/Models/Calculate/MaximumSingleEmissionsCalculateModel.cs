using System.ComponentModel.DataAnnotations;
using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models.Calculate;

public class MaximumSingleEmissionsCalculateModel
{
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
    [Range(15, 25)]
    public float AvgExitSpeed { get; set; }

    /// <summary>
    /// Высота источника выброса, м.
    /// </summary>
    [Range(13, 65)]
    public float HeightSource { get; set; }

    /// <summary>
    /// Диаметр устья источника, м.
    /// </summary>
    [Range(1, 7)]
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
    public int Distance { get; set; }
}