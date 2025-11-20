namespace EcoPulseBackend.Models;

/// <summary>
/// Отчет по расчету выбросов ЗВ при сварочных работах (входные + выходные данные)
/// </summary>
public class DuringWeldingOperationsEmissionsReport
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
    /// Расход сварочных электродов в год, кг
    /// </summary>
    public float ElectrodesPerYear { get; set; }
    
    /// <summary>
    /// Время работы сварочного оборудования, ч/год
    /// </summary>
    public int WorkDaysPerYear { get; set; }
    
    /// <summary>
    /// Результат расчетов выбросов ЗВ при сварочных работах
    /// </summary>
    public DuringWeldingOperationsEmissionsBatchResult Result { get; set; } = new();
}