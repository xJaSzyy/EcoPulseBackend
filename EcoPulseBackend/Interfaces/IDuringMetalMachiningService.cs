using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DuringMetalMachining;

namespace EcoPulseBackend.Interfaces;

/// <summary>
/// Сервис механической обработки металлов
/// </summary>
public interface IDuringMetalMachiningService
{
    /// <summary>
    /// Расчет выбросов по группе загрязнителей
    /// </summary>
    /// <param name="model">Модель для расчета выбросов при механической обработке металлов</param>
    /// <returns></returns>
    public List<EmissionsResult> CalculateEmissionsBatch(DuringMetalMachiningEmissionsCalculateModel model);
}