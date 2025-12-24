using EcoPulseBackend.Models.DuringWeldingOperations;

namespace EcoPulseBackend.Interfaces;

public interface IDuringWeldingOperationsService
{
    /// <summary>
    /// Расчет выбросов загрязняющих веществ при сварочных работах
    /// </summary>
    /// <param name="model">Модель для расчета выбросов загрязняющих веществ при сварочных работах</param>
    /// <returns></returns>
    public DuringWeldingOperationsEmissionsBatchResult CalculateDuringWeldingOperationsEmissionsBatch(DuringWeldingOperationsEmissionsCalculateModel model);
}