namespace EcoPulseBackend.Models.OpenCoalWarehouse;

/// <summary>
/// Модель для расчета выбросов угольной пыли в атмосферу от открытых складов угля
/// </summary>
public class OpenCoalWarehouseEmissionsCalculateModel
{
    /// <summary>
    /// Удельное выделение твердых частиц при разгрузке материала, г/т
    /// </summary>
    public float SpecificEmission { get; set; }
    
    /// <summary>
    /// Количество разгружаемого материала, т/г
    /// </summary>
    public float UnloadMaterialCountPerYear { get; set; }
    
    /// <summary>
    /// Количество разгружаемого материала, т/ч
    /// </summary>
    public float UnloadMaterialCountPerHour { get; set; }
    
    /// <summary>
    /// Эффективность применяемого средства пылеподавления, дол. ед.
    /// </summary>
    public float DustSuppressionEfficiency { get; set; }
    
    /// <summary>
    /// Площадь основания штабеля угля, м2
    /// </summary>
    public float CoalPileBaseArea { get; set; }
        
    /// <summary>
    /// Количество дней с устойчивым снежным покровом
    /// </summary>
    public int SnowyDaysCount { get; set; }
    
    /// <summary>
    /// Количество дней с осадками в виде дождя
    /// </summary>
    public int RainyDaysCount { get; set; } 
}