using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models;

/// <summary>
/// Отчет по расчету выбросов ЗВ от резервуаров (входные + выходные данные)
/// </summary>
public class ReservoirsEmissionsReport
{
    /// <summary>
    /// Номер источника выделения
    /// </summary>
    public string SelectionSource { get; set; } = null!;

    /// <summary>
    /// Номер источника загрязнения
    /// </summary>
    public string PollutionSource { get; set; } = null!;
    
    /// <summary>
    /// Объем резервуара, м3
    /// </summary>
    public float ReservoirVolume { get; set; }
    
    /// <summary>
    /// Количество резервуаров, шт.
    /// </summary>
    public float ReservoirCount { get; set; }
    
    /// <summary>
    /// Время работы, ч/г
    /// </summary>
    public float WorkHoursPerYear { get; set; }
    
    /// <summary>
    /// Конструкция резервуара
    /// </summary>
    public ReservoirType ReservoirType { get; set; }
    
    /// <summary>
    /// Нефтепродукт
    /// </summary>
    public OilProduct OilProduct { get; set; }
    
    /// <summary>
    /// Климатическая зона
    /// </summary>
    public ClimateZone ClimateZone { get; set; }
    
    /// <summary>
    /// Концентрация паров нефтепродуктов в выбросах при заполнении резервуаров
    /// </summary>
    public VaporConcentrationRecord VaporConcentration { get; set; } = null!;

    /// <summary>
    /// Кол-во закачиваемого в резервуар нефтепродукта в осенне-зимний период, м3
    /// </summary>
    public float AutumnWinterOilAmount { get; set; }
    
    /// <summary>
    /// Кол-во закачиваемого в резервуар нефтепродукта в весенне-летний период, м3
    /// </summary>
    public float SpringSummerOilAmount { get; set; }
    
    /// <summary>
    /// Объем слитого нефтепродукта в резервуар, м3
    /// </summary>
    public float DrainedVolume { get; set; }
    
    /// <summary>
    /// Среднее время слива, с
    /// </summary>
    public float AverageDrainTime { get; set; }
    
    /// <summary>
    /// Результат расчетов выбросов ЗВ от резервуаров
    /// </summary>
    public ReservoirsEmissionsBatchResult Result { get; set; } = new();
}