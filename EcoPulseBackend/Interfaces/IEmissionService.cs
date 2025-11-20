using EcoPulseBackend.Enums;
using EcoPulseBackend.Models;

namespace EcoPulseBackend.Interfaces;

public interface IEmissionService
{
    /// <summary>
    /// Расчет выбросов загрязняющих веществ от бензогенератора
    /// </summary>
    /// <param name="pollutants">Список загрязняющих веществ</param>
    /// <param name="workHoursPerDay">Время работы в день, ч</param>
    /// <param name="workDaysPerYear">Кол-во рабочих дней в году</param>
    /// <param name="generatorCount">Кол-во генераторов, шт</param>
    /// <param name="sameGeneratorCount">Кол-во одновременно работающих генераторов, шт</param>
    public List<EmissionsResult> CalculateGasolineGeneratorEmissionsBatch(List<Pollutant> pollutants,
        int workHoursPerDay, int workDaysPerYear, int generatorCount, int sameGeneratorCount);

    /// <summary>
    /// Расчет выбросов загрязняющих веществ от резервуаров
    /// </summary>
    /// <param name="pollutants">Список загрязняющих веществ</param>
    /// <param name="vaporConcentration">Концентрация паров нефтепродуктов в выбросах при заполнении резервуаров</param>
    /// <param name="autumnWinterOilAmount">Кол-во закачиваемого в резервуар нефтепродукта в осенне-зимний период, м3</param>
    /// <param name="springSummerOilAmount">Кол-во закачиваемого в резервуар нефтепродукта в весенне-летний период, м3</param>
    /// <param name="drainedVolume">Объем слитого нефтепродукта в резервуар, м3</param>
    /// <param name="averageDrainTime">Среднее время слива, с</param>
    /// <returns></returns>
    public ReservoirsEmissionsBatchResult CalculateReservoirsEmissionsBatch(List<Pollutant> pollutants,
        VaporConcentrationRecord vaporConcentration, float autumnWinterOilAmount, float springSummerOilAmount,
        float drainedVolume, float averageDrainTime = 1200f);

    /// <summary>
    /// Расчет выбросов загрязняющих веществ при механической обработке металлов
    /// </summary>
    /// <param name="metalMachiningMachineType">Тип станка для обработки металла</param>
    /// <param name="workDaysPerYear">Годовой фонд времени работы оборудования, ч</param>
    /// <returns></returns>
    public EmissionsResult CalculateDuringMetalMachiningEmissions(MetalMachiningMachineType metalMachiningMachineType, int workDaysPerYear);

    /// <summary>
    /// Расчет выбросов загрязняющих веществ при сварочных работах
    /// </summary>
    /// <param name="pollutants">Список загрязняющих веществ</param>
    /// <param name="electrodesPerYear">Расход сварочных электродов в год, кг</param>
    /// <param name="workDaysPerYear">Время работы сварочного оборудования, ч/год (Т)</param>
    /// <returns></returns>
    public DuringWeldingOperationsEmissionsBatchResult CalculateDuringWeldingOperationsEmissionsBatch(
        List<Pollutant> pollutants,
        float electrodesPerYear, int workDaysPerYear);
}