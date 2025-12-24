using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DangerZone;
using EcoPulseBackend.Models.TrafficLightQueue;
using EcoPulseBackend.Models.TrafficLightQueueEmissionSource;

namespace EcoPulseBackend.Interfaces;

/// <summary>
/// Сервис стоящего транспорта на регулируемом перекрестке
/// </summary>
public interface ITrafficLightQueueService
{
    /// <summary>
    /// Расчет выбросов по группе загрязнителей
    /// </summary>
    /// <param name="model">Модель для расчета выбросов стоящего транспорта на регулируемом перекрестке</param>
    /// <returns></returns>
    public List<EmissionsResult> CalculateEmissionsBatch(TrafficLightQueueEmissionsCalculateModel model);
    
    /// <summary>
    /// Расчет нескольких зон выбросов
    /// </summary>
    /// <param name="emissionSources">Список источников выбросов</param>
    /// <returns></returns>
    public List<TrafficLightQueueDangerZone> CalculateDangerZones(List<TrafficLightQueueEmissionSource> emissionSources);
}