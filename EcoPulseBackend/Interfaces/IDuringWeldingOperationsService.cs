using EcoPulseBackend.Models.DuringWeldingOperations;

namespace EcoPulseBackend.Interfaces;

/// <summary>
/// Сервис сварочных работ
/// </summary>
public interface IDuringWeldingOperationsService
{
    /// <summary>
    /// Расчет выбросов по группе загрязнителей
    /// </summary>
    /// <param name="model">Модель для расчета выбросов при сварочных работах</param>
    /// <returns></returns>
    public DuringWeldingOperationsEmissionsBatchResult CalculateEmissionsBatch(DuringWeldingOperationsEmissionsCalculateModel model);
}