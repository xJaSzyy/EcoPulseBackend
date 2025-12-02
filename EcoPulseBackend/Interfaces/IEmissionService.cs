using EcoPulseBackend.Enums;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.Calculate;

namespace EcoPulseBackend.Interfaces;

public interface IEmissionService
{
    /// <summary>
    /// Расчет выбросов загрязняющих веществ от бензогенератора
    /// </summary>
    /// <param name="model">Модель для расчета выбросов загрязняющих веществ от бензогенератора</param>
    public List<EmissionsResult> CalculateGasolineGeneratorEmissionsBatch(GasolineGeneratorEmissionsCalculateModel model);

    /// <summary>
    /// Расчет выбросов загрязняющих веществ от резервуаров
    /// </summary>
    /// <param name="model">Модель для расчета выбросов загрязняющих веществ от резервуаров</param>
    /// <returns></returns>
    public ReservoirsEmissionsBatchResult CalculateReservoirsEmissionsBatch(ReservoirsEmissionsCalculateModel model);

    /// <summary>
    /// Расчет выбросов загрязняющих веществ при механической обработке металлов
    /// </summary>
    /// <param name="model">Модель для расчета выбросов загрязняющих веществ при механической обработке металлов</param>
    /// <returns></returns>
    public List<EmissionsResult> CalculateDuringMetalMachiningEmissionsBatch(DuringMetalMachiningEmissionsCalculateModel model);

    /// <summary>
    /// Расчет выбросов загрязняющих веществ при сварочных работах
    /// </summary>
    /// <param name="model">Модель для расчета выбросов загрязняющих веществ при сварочных работах</param>
    /// <returns></returns>
    public DuringWeldingOperationsEmissionsBatchResult CalculateDuringWeldingOperationsEmissionsBatch(DuringWeldingOperationsEmissionsCalculateModel model);

    /// <summary>
    /// Расчет выбросов загрязняющих вещество от одиночного точечного источника
    /// </summary>
    /// <param name="model">Модель для расчета выбросов загрязняющих вещество от одиночного точечного источника</param>
    /// <returns></returns>
    public EmissionsGroupResult CalculateMaximumSingleEmissions(MaximumSingleEmissionsCalculateModel model);

    /// <summary>
    /// Расчет выбросов движущегося автотранспорта
    /// </summary>
    /// <param name="model">Модель для расчета выбросов движущегося автотранспорта</param>
    /// <returns></returns>
    public List<EmissionsResult> CalculateVehicleFlowEmissionsBatch(VehicleFlowEmissionsCalculateModel model);

    /// <summary>
    /// Расчет выбросов автотранспорта в районе регулируемого перекрестка
    /// </summary>
    /// <param name="model">Модель для расчета выбросов автотранспорта в районе регулируемого перекрестка</param>
    /// <returns></returns>
    public List<EmissionsResult> CalculateTrafficLightQueueEmissionsBatch(TrafficLightQueueEmissionsCalculateModel model);


    /// <summary>
    /// Расчет выбросов угольной пыли в атмосферу от открытых складов угля
    /// </summary>
    /// <param name="model">Модель для расчета выбросов угольной пыли в атмосферу от открытых складов угля</param>
    /// <returns></returns>
    public EmissionsResult CalculateOpenCoalWarehouseEmissions(OpenCoalWarehouseEmissionsCalculateModel model);
}