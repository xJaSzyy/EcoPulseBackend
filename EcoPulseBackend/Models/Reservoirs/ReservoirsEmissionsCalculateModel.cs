using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models.Reservoirs;

/// <summary>
/// Модель для расчета выбросов загрязняющих веществ от резервуаров
/// </summary>
public class ReservoirsEmissionsCalculateModel
{
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
}