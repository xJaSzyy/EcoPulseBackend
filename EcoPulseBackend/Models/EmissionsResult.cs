namespace EcoPulseBackend.Models;

/// <summary>
/// Результат одного расчета выбросов ЗВ
/// </summary>
public class EmissionsResult
{
    public PollutantInfo PollutantInfo { get; set; } = null!;

    public float MaximumEmission { get; set; }

    public float GrossEmission { get; set; }
}