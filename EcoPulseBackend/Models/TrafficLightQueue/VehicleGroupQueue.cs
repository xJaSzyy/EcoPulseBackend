using System.ComponentModel.DataAnnotations;
using EcoPulseBackend.Enums;

namespace EcoPulseBackend.Models.TrafficLightQueue;

public class VehicleGroupQueue
{
    /// <summary>
    /// Идентификатор
    /// </summary>
    [Key]
    public int Id { get; set; }
    
    /// <summary>
    /// Идентификатор источника выброса
    /// </summary>
    public int TrafficLightQueueEmissionSourceId { get; set; }
    
    /// <summary>
    /// Тип транспортного средства
    /// </summary>
    public VehicleType VehicleType { get; set; }
        
    /// <summary>
    /// Количество автомобилей, находящихся в «очереди» в зоне перекрестка в конце п-го цикла запрещающего сигнала светофора
    /// </summary>
    public int VehiclesCount { get; set; }
}