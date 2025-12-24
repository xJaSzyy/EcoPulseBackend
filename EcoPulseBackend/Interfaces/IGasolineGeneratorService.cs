using EcoPulseBackend.Models;
using EcoPulseBackend.Models.GasolineGenerator;

namespace EcoPulseBackend.Interfaces;

public interface IGasolineGeneratorService
{
    /// <summary>
    /// Расчет выбросов загрязняющих веществ от бензогенератора
    /// </summary>
    /// <param name="model">Модель для расчета выбросов загрязняющих веществ от бензогенератора</param>
    public List<EmissionsResult> CalculateGasolineGeneratorEmissionsBatch(GasolineGeneratorEmissionsCalculateModel model);
}