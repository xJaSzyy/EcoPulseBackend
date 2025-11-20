using System.ComponentModel;

namespace EcoPulseBackend.Enums;

/// <summary>
/// Климатическая зона
/// </summary>
public enum ClimateZone
{
    /// <summary>
    /// 1-я климатическая зона
    /// </summary>
    [Description("1-я климатическая зона")]
    First = 1,
    
    /// <summary>
    /// 2-я климатическая зона
    /// </summary>
    [Description("2-я климатическая зона")]
    Second = 2,
    
    /// <summary>
    /// 3-я климатическая зона
    /// </summary>
    [Description("3-я климатическая зона")]
    Third = 3
}