namespace EcoPulseBackend.Models.Calculate;

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