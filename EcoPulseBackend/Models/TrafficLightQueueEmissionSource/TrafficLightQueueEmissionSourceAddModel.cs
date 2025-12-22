using EcoPulseBackend.Enums;
using EcoPulseBackend.Models.TrafficLightQueue;

namespace EcoPulseBackend.Models.TrafficLightQueueEmissionSource;

public class TrafficLightQueueEmissionSourceAddModel
{
    /// <summary>
    /// Начальные координаты
    /// </summary>
    public Coordinates Location { get; set; } = null!;
    
    /// <summary>
    /// Список групп транспортных средств, стоящих в очереди
    /// </summary>
    public List<VehicleGroupQueue> VehicleGroups { get; set; } = [];
    
    /// <summary>
    /// Количество циклов действия запрещающего сигнала светофора за 20-минутный период времени
    /// </summary>
    public int TrafficLightCycles { get; set; }
    
    /// <summary>
    /// Продолжительность действия запрещающего сигнала светофора (включая желтый цвет)
    /// </summary>
    public float TrafficLightStopTime { get; set; }
}