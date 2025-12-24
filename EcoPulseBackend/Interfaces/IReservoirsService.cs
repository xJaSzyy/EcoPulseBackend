using EcoPulseBackend.Models.Reservoirs;

namespace EcoPulseBackend.Interfaces;

public interface IReservoirsService
{
    /// <summary>
    /// Расчет выбросов загрязняющих веществ от резервуаров
    /// </summary>
    /// <param name="model">Модель для расчета выбросов загрязняющих веществ от резервуаров</param>
    /// <returns></returns>
    public ReservoirsEmissionsBatchResult CalculateReservoirsEmissionsBatch(ReservoirsEmissionsCalculateModel model);
}