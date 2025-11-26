using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models.Calculate;

public class DuringMetalMachiningEmissionsCalculateModel
{
    /// <summary>
    /// Тип станка для обработки металла
    /// </summary>
    public MetalMachiningMachineType MetalMachiningMachineType { get; set; }
    
    /// <summary>
    /// Годовой фонд времени работы оборудования, ч
    /// </summary>
    public int WorkDaysPerYear { get; set; }
}