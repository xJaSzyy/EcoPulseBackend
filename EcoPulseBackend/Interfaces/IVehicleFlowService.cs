using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DangerZone;
using EcoPulseBackend.Models.VehicleFlow;
using EcoPulseBackend.Models.VehicleFlowEmissionSource;

namespace EcoPulseBackend.Interfaces;

public interface IVehicleFlowService
{
    /// <summary>
    /// Расчет выбросов движущегося автотранспорта
    /// </summary>
    /// <param name="model">Модель для расчета выбросов движущегося автотранспорта</param>
    /// <returns></returns>
    public List<EmissionsResult> CalculateVehicleFlowEmissionsBatch(VehicleFlowEmissionsCalculateModel model);
    
    public List<VehicleFlowDangerZone> CalculateVehicleFlowEmissionDangerZones(List<VehicleFlowEmissionSource> emissionSources);
}