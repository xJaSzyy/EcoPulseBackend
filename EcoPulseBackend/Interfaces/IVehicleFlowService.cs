using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DangerZone;
using EcoPulseBackend.Models.VehicleFlow;
using EcoPulseBackend.Models.VehicleFlowEmissionSource;

namespace EcoPulseBackend.Interfaces;

/// <summary>
/// Сервис движущегося транспорта
/// </summary>
public interface IVehicleFlowService
{
    /// <summary>
    /// Расчет выбросов по группе загрязнителей
    /// </summary>
    /// <param name="model">Модель для расчета выбросов движущегося транспорта</param>
    /// <returns></returns>
    public List<EmissionsResult> CalculateEmissionsBatch(VehicleFlowEmissionsCalculateModel model);
    
    /// <summary>
    /// Расчет нескольких зон выбросов
    /// </summary>
    /// <param name="emissionSources">Список источников выбросов</param>
    /// <returns></returns>
    public List<VehicleFlowDangerZone> CalculateDangerZones(List<VehicleFlowEmissionSource> emissionSources);
}