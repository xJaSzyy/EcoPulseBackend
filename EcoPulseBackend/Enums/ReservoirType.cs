using System.ComponentModel;

namespace EcoPulseBackend.Enums;

/// <summary>
/// Конструкция резервуара
/// </summary>
public enum ReservoirType
{
    /// <summary>
    /// Наземный
    /// </summary>
    [Description("Наземный")]
    Ground = 1,
    
    /// <summary>
    /// Заглубленный
    /// </summary>
    [Description("Заглубленный")]
    Buried = 2
}