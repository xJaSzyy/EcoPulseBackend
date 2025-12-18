using System.ComponentModel.DataAnnotations;
using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models.SingleEmissionSource;

/// <summary>
/// Модель изменения источника выброса
/// </summary>
public class SingleEmissionSourceUpdateModel
{
    /// <summary>
    /// Идентификатор
    /// </summary>
    public int Id { get; set; }
    
    /// <summary>
    /// Долгота
    /// </summary>
    public float Lon { get; set; }
    
    /// <summary>
    /// Широта
    /// </summary>
    public float Lat { get; set; }
    
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