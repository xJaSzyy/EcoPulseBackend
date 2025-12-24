using EcoPulseBackend.Models;
using EcoPulseBackend.Models.OpenCoalWarehouse;

namespace EcoPulseBackend.Interfaces;

public interface IOpenCoalWarehouseService
{
    /// <summary>
    /// Расчет выбросов угольной пыли в атмосферу от открытых складов угля
    /// </summary>
    /// <param name="model">Модель для расчета выбросов угольной пыли в атмосферу от открытых складов угля</param>
    /// <returns></returns>
    public List<EmissionsResult> CalculateOpenCoalWarehouseEmissions(OpenCoalWarehouseEmissionsCalculateModel model);
}