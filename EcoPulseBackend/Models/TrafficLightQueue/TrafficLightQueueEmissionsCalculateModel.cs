namespace EcoPulseBackend.Models.TrafficLightQueue;

/// <summary>
/// Модель для расчета выбросов автотранспорта в районе регулируемого перекрестка
/// </summary>
public class TrafficLightQueueEmissionsCalculateModel
{
    /// <summary>
    /// Список групп транспортных средств, стоящих в очереди
    /// </summary>
    public List<VehicleGroupQueue> VehicleGroups { get; set; } = new();
    
    /// <summary>
    /// Количество циклов действия запрещающего сигнала светофора за 20-минутный период времени
    /// </summary>
    public int TrafficLightCycles { get; set; }
    
    /// <summary>
    /// Продолжительность действия запрещающего сигнала светофора (включая желтый цвет)
    /// </summary>
    public float TrafficLightStopTime { get; set; }
}