using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models;

/// <summary>
/// Отчет по расчету выбросов ЗВ при механической обработке металлов (входные + выходные данные)
/// </summary>
public class DuringMetalMachiningEmissionsReport
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
    /// Тип станка для обработки металла
    /// </summary>
    public MetalMachiningMachineType MetalMachiningMachineType { get; set; }
    
    /// <summary>
    /// Годовой фонд времени работы оборудования, ч
    /// </summary>
    public int WorkDaysPerYear { get; set; }
    
    /// <summary>
    /// Число оборудования данного типа (n)
    /// </summary>
    public int MachiningMachineCount { get; set; }
    
    /// <summary>
    /// Число оборудования данного типа работающего одновременно (n)
    /// </summary>
    public int SameMachiningMachineCount { get; set; }
    
    /// <summary>
    /// Результат расчетов выбросов ЗВ при механической обработке металлов
    /// </summary>
    public EmissionsResult Result { get; set; } = new();
}