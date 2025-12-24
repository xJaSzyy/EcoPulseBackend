using EcoPulseBackend.Models;
using EcoPulseBackend.Models.OpenCoalWarehouse;

namespace EcoPulseBackend.Interfaces;

/// <summary>
/// Сервис открытого склада угля
/// </summary>
public interface IOpenCoalWarehouseService
{
    /// <summary>
    /// Расчет выбросов по одному загрязнителю
    /// </summary>
    /// <param name="model">Модель для расчета выбросов угольной пыли от открытых складов угля</param>
    /// <returns></returns>
    public List<EmissionsResult> CalculateEmissions(OpenCoalWarehouseEmissionsCalculateModel model);
}