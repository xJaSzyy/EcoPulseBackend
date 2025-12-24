using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DangerZone;
using EcoPulseBackend.Models.MaximumSingle;

namespace EcoPulseBackend.Interfaces;

public interface IMaximumSingleService
{
    /// <summary>
    /// Расчет выбросов загрязняющих вещество от одиночного точечного источника
    /// </summary>
    /// <param name="model">Модель для расчета выбросов загрязняющих вещество от одиночного точечного источника</param>
    /// <returns></returns>
    public EmissionsGroupResult CalculateMaximumSingleEmissions(MaximumSingleEmissionsCalculateModel model);
    
    /// <summary>
    /// Расчет зоны выброса от одиночного точечного источника
    /// </summary>
    /// <param name="model">Модель для расчета выбросов загрязняющих вещество от одиночного точечного источника</param>
    /// <returns></returns>
    public SingleDangerZone CalculateMaximumSingleEmissionsDangerZone(MaximumSingleEmissionsCalculateModel model);
}