namespace EcoPulseBackend.Models.DangerZone;

public class TrafficLightQueueDangerZoneCalculateModel
{
    /// <summary>
    /// Список идентификаторов городов
    /// </summary>
    public ICollection<int> CityIds { get; set; } = new List<int>();
}