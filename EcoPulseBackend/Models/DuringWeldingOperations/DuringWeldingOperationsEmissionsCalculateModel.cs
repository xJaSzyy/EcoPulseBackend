namespace EcoPulseBackend.Models.DuringWeldingOperations;

/// <summary>
/// Модель для расчета выбросов загрязняющих веществ при сварочных работах
/// </summary>
public class DuringWeldingOperationsEmissionsCalculateModel
{
    /// <summary>
    /// Расход сварочных электродов в год, кг
    /// </summary>
    public float ElectrodesPerYear { get; set; }
    
    /// <summary>
    /// Время работы сварочного оборудования, ч/год
    /// </summary>
    public int WorkDaysPerYear { get; set; }
}