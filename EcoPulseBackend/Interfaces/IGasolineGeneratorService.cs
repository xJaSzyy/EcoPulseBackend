using EcoPulseBackend.Models;
using EcoPulseBackend.Models.GasolineGenerator;

namespace EcoPulseBackend.Interfaces;

/// <summary>
/// Сервис бензогенератора
/// </summary>
public interface IGasolineGeneratorService
{
    /// <summary>
    /// Расчет выбросов по группе загрязнителей
    /// </summary>
    /// <param name="model">Модель для расчета выбросов от бензогенератора</param>
    public List<EmissionsResult> CalculateEmissionsBatch(GasolineGeneratorEmissionsCalculateModel model);
}