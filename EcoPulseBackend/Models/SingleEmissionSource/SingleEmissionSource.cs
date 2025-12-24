using System.ComponentModel.DataAnnotations;
using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models.SingleEmissionSource;

/// <summary>
/// Одиночный источник выброса
/// </summary>
public class SingleEmissionSource
{
    /// <summary>
    /// Идентификатор
    /// </summary>
    [Key]
    public int Id { get; set; }
    
    /// <summary>
    /// Идентификатор города
    /// </summary>
    public int CityId { get; set; }
    
    /// <summary>
    /// Город
    /// </summary>
    public City City { get; set; } = null!;
    
    /// <summary>
    /// Координаты
    /// </summary>
    public Coordinates Location { get; set; } = null!;

    /// <summary>
    /// Температура выбрасываемой ГВС
    /// </summary>
    [Range(235, 265)]
    public float EjectedTemp { get; set; }

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
}