using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models;

public class VehicleGroupQueue
{
    /// <summary>
    /// Тип транспортного средства
    /// </summary>
    public VehicleType VehicleType { get; set; }
        
    /// <summary>
    /// Количество автомобилей, находящихся в «очереди» в зоне перекрестка в конце п-го цикла запрещающего сигнала светофора
    /// </summary>
    public int VehiclesCount { get; set; }
        
    /// <summary>
    /// Средняя скорость движения транспортного потока
    /// </summary>
    public float AverageSpeed { get; set; }
}