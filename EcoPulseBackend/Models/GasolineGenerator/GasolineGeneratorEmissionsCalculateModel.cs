namespace EcoPulseBackend.Models.GasolineGenerator;

/// <summary>
/// Модель для расчета выбросов загрязняющих веществ от бензогенератора
/// </summary>
public class GasolineGeneratorEmissionsCalculateModel
{
    /// <summary>
    /// Время работы в день, ч
    /// </summary>
    public int WorkHoursPerDay { get; set; }

    /// <summary>
    /// Количество рабочих дней в году
    /// </summary>
    public int WorkDaysPerYear { get; set; }

    /// <summary>
    /// Количество генераторов, k-вида
    /// </summary>
    public int GeneratorCount { get; set; }

    /// <summary>
    /// Количество одновременно работающих генераторов k-вида
    /// </summary>
    public int SameGeneratorCount { get; set; }
}