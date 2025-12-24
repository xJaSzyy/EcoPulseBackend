using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DuringMetalMachining;

namespace EcoPulseBackend.Interfaces;

public interface IDuringMetalMachiningService
{
    /// <summary>
    /// Расчет выбросов загрязняющих веществ при механической обработке металлов
    /// </summary>
    /// <param name="model">Модель для расчета выбросов загрязняющих веществ при механической обработке металлов</param>
    /// <returns></returns>
    public List<EmissionsResult> CalculateDuringMetalMachiningEmissionsBatch(DuringMetalMachiningEmissionsCalculateModel model);
}