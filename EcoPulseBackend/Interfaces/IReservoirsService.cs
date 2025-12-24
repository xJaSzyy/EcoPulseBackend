using EcoPulseBackend.Models.Reservoirs;

namespace EcoPulseBackend.Interfaces;

/// <summary>
/// Сервис резервуаров
/// </summary>
public interface IReservoirsService
{
    /// <summary>
    /// Расчет выбросов по группе загрязнителей
    /// </summary>
    /// <param name="model">Модель для расчета выбросов от резервуаров</param>
    /// <returns></returns>
    public ReservoirsEmissionsBatchResult CalculateEmissionsBatch(ReservoirsEmissionsCalculateModel model);
}