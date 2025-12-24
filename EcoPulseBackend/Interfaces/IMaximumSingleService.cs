using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DangerZone;
using EcoPulseBackend.Models.MaximumSingle;

namespace EcoPulseBackend.Interfaces;

/// <summary>
/// Сервис одиночного точечного источника
/// </summary>
public interface IMaximumSingleService
{
    /// <summary>
    /// Расчет выбросов по одному загрязнителю
    /// </summary>
    /// <param name="model">Модель для расчета выбросов от одиночного точечного источника</param>
    /// <returns></returns>
    public EmissionsGroupResult CalculateEmissions(MaximumSingleEmissionsCalculateModel model);
    
    /// <summary>
    /// Расчет зоны выброса
    /// </summary>
    /// <param name="model">Модель для расчета выбросов от одиночного точечного источника</param>
    /// <returns></returns>
    public SingleDangerZone CalculateDangerZone(MaximumSingleEmissionsCalculateModel model);
}