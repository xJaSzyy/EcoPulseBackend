using System.Text.Json.Serialization;

namespace EcoPulseBackend.Models;

/// <summary>
/// Отчет по расчету выбросов ЗВ от бензогенератора (входные + выходные данные)
/// </summary>
public class GasolineGeneratorEmissionsReport
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

    /// <summary>
    /// Список результатов расчетов выбросов ЗВ от бензогенератора
    /// </summary>
    [JsonIgnore]
    public List<EmissionsResult> Emissions { get; set; } = new();
}