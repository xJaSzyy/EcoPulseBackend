using System.ComponentModel;

namespace EcoPulseBackend.Enums;

/// <summary>
/// Тип станка для обработки металла
/// </summary>
public enum MetalMachiningMachineType
{
    /// <summary>
    /// Сверлильный
    /// </summary>
    [Description("Сверлильный станок")]
    Drilling = 1, 
    
    /// <summary>
    /// Крацевальный
    /// </summary>
    [Description("Крацевальный станок")]
    Milling = 2,
    
    /// <summary>
    /// Отрезной
    /// </summary>
    [Description("Отрезной станок")]
    Cutting = 3
}