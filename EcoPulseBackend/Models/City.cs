using System.ComponentModel.DataAnnotations;

namespace EcoPulseBackend.Models;

/// <summary>
/// Город
/// </summary>
public class City
{
    /// <summary>
    /// Идентификатор
    /// </summary>
    [Key]
    public int Id { get; set; }
    
    /// <summary>
    /// Название
    /// </summary>
    public string Name { get; set; } = null!;
    
    /// <summary>
    /// Координаты
    /// </summary>
    public Coordinates Location { get; set; } = null!;
}