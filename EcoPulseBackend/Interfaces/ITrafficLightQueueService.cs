using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DangerZone;
using EcoPulseBackend.Models.TrafficLightQueue;
using EcoPulseBackend.Models.TrafficLightQueueEmissionSource;

namespace EcoPulseBackend.Interfaces;

public interface ITrafficLightQueueService
{
    /// <summary>
    /// Расчет выбросов автотранспорта в районе регулируемого перекрестка
    /// </summary>
    /// <param name="model">Модель для расчета выбросов автотранспорта в районе регулируемого перекрестка</param>
    /// <returns></returns>
    public List<EmissionsResult> CalculateTrafficLightQueueEmissionsBatch(TrafficLightQueueEmissionsCalculateModel model);
    
    public List<TrafficLightQueueDangerZone> CalculateTrafficLightQueueEmissionDangerZones(List<TrafficLightQueueEmissionSource> emissionSources);
}