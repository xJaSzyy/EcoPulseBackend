using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models.TrafficLightQueueEmissionSource;

public class TrafficLightQueueEmissionSourceAddModel
{
    /// <summary>
    /// Начальные координаты
    /// </summary>
    public Coordinates Location { get; set; } = null!;
    
    /// <summary>
    /// Тип транспортного средства
    /// </summary>
    public VehicleType VehicleType { get; set; }
        
    /// <summary>
    /// Количество автомобилей, находящихся в «очереди» в зоне перекрестка в конце п-го цикла запрещающего сигнала светофора
    /// </summary>
    public int VehiclesCount { get; set; }
    
    /// <summary>
    /// Количество циклов действия запрещающего сигнала светофора за 20-минутный период времени
    /// </summary>
    public int TrafficLightCycles { get; set; }
    
    /// <summary>
    /// Продолжительность действия запрещающего сигнала светофора (включая желтый цвет)
    /// </summary>
    public float TrafficLightStopTime { get; set; }
}