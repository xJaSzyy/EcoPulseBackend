using ClosedXML.Excel;
using EcoPulseBackend.Extensions;
using EcoPulseBackend.Interfaces;
using EcoPulseBackend.Models;

namespace EcoPulseBackend.Services;

public class ExportService : IExportService
{
    public MemoryStream CreateGasolineGeneratorEmissionsReport(GasolineGeneratorEmissionsReport report)
    {
        var sheetName = $"ИЗА {report.PollutionSource}_{report.SelectionSource}";
        
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(sheetName);

        var row = 1;
        
        SetCell(worksheet, $"A{row}:I{row}", "Расчет выбросов загрязняющих веществ от бензогенератора", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"A{row + 2}:I{row + 2}", $"Источник загрязнения № {report.PollutionSource}, Бензогенератор", true);
        SetCell(worksheet, $"A{row + 3}:I{row + 3}", $"Источник выделения № {report.SelectionSource}", true);
        SetCell(worksheet, $"A{row + 5}:I{row + 5}", "Литература: Методика проведения инвентаризации выбросов загрязняющих веществ в атмосферу автотранспортных предприятий (расчетным методом). Москва, 1998, с дополнениями и изменениями к Методике проведения инвентаризации выбросов загрязняющих веществ в атмосферу автотранспортных предприятий (расчетным методом). М, 1999");
        SetCell(worksheet, $"A{row + 7}:I{row + 7}", "В соответствии с п.12 \"Методического пособия по расчету, нормированию и контролю выбросов загрязняющих веществ в атмосферный воздух\", СПб., 2012 расчет выбросов от бензогенераторов выполняем по \"Методике проведения инвентаризации выбросов загрязняющих веществ в атмосферу автотранспортных предприятий (расчетным методом). Москва, 1998, с дополнениями и изменениями к Методике проведения инвентаризации выбросов загрязняющих веществ в атмосферу автотранспортных предприятий (расчетным методом). М, 1999\", принимая за выброс - 0,25 от величины выброса легкового карбюраторнорго автомобиля с объемом двигателя  до 1,2 л при движении по территории со скоростью 5 км/ч.");
        row += 9;
        
        SetCell(worksheet, $"A{row}:I{row}", "Валовый выброс определяется по формуле:", true);
        SetCell(worksheet, $"A{row + 2}:I{row + 2}", "Mi = 0,25 \u00d7 gi \u00d7 5,0 \u00d7 ti \u00d7 b \u00d7 Nk  / 1000000, т/г", true);
        SetCell(worksheet, $"A{row + 4}:I{row + 4}", "где  gi - удельный выброс, г/км (удельные выбросы - пробеговые выбросы, г/км) (табл. 2.5)");
        SetCell(worksheet, $"A{row + 5}:I{row + 5}", "ti - время работы в день, ч;");
        SetCell(worksheet, $"A{row + 6}:I{row + 6}", "b - количество рабочих дней в году;");
        SetCell(worksheet, $"A{row + 7}:I{row + 7}", "Nk - количество генераторов, k-вида, шт;");
        SetCell(worksheet, $"A{row + 8}:I{row + 8}", "5.0 - скорость движения км/ч;");
        SetCell(worksheet, $"A{row + 9}:I{row + 9}", "1000000 - перевод г на тонны.");
        row += 11;
        
        SetCell(worksheet, $"A{row}:I{row}", "Максимально разовый выброс определяется по формуле:", true);
        SetCell(worksheet, $"A{row + 2}:I{row + 2}", "Gi = 0,25 \u00d7 gi \u00d7 5 \u00d7 nk / 3600, г/с", true);
        SetCell(worksheet, $"A{row + 4}:I{row + 4}", "где nk - количество одновременно работающих генераторов k-вида;");
        SetCell(worksheet, $"A{row + 5}:I{row + 5}", "3600 - перевод г/ч на г/с.");
        row += 7;
        
        SetCell(worksheet, $"A{row}:I{row}", "Расчет выбросов", true);
        SetCell(worksheet, $"A{row + 1}:A{row + 2}", "Наименование генератора", horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"B{row + 1}:B{row + 2}", "Кол-во, nk, шт.", horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"C{row + 1}:C{row + 2}", "Кол-во, Nk, шт.", horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"D{row + 1}:D{row + 2}", "Время работы в день, ti ,ч", horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"E{row + 1}:E{row + 2}", "Кол-во рабочих дней в год, b", horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"F{row + 1}:F{row + 2}", "Наименование ЗВ", horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"G{row + 1}:G{row + 2}", "Удельный выброс, gi, г/км", horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"H{row + 1}:I{row + 1}", "Выбросы в атмосферу", horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"H{row + 2}", "Максимально-разовый выброс, г/с", horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"I{row + 2}", "Валовый выброс, т/г", horizontal: XLAlignmentHorizontalValues.Center);
        row += 3;
        
        SetCell(worksheet, $"A{row}", 1, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"B{row}", 2, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"C{row}", 3, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"D{row}", 4, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"E{row}", 5, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"F{row}", 6, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"G{row}", 7, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"H{row}", 8, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"I{row}", 9, horizontal: XLAlignmentHorizontalValues.Center);
        
        SetCell(worksheet, $"A{row + 1}:A{row + 5}", "Бензиновый генератор", horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"B{row + 1}:B{row + 5}", report.SameGeneratorCount, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"C{row + 1}:C{row + 5}", report.GeneratorCount, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"D{row + 1}:D{row + 5}", report.WorkHoursPerDay, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"E{row + 1}:E{row + 5}", report.WorkDaysPerYear, horizontal: XLAlignmentHorizontalValues.Center);
        
        foreach (var emission in report.Emissions)
        {
            row++;
            SetCell(worksheet, $"F{row}", emission.PollutantInfo.Pollutant.ToString(), horizontal: XLAlignmentHorizontalValues.Center);
            SetCell(worksheet, $"G{row}", emission.PollutantInfo.SpecificEmission, horizontal: XLAlignmentHorizontalValues.Center);
            SetCell(worksheet, $"H{row}", emission.MaximumEmission, horizontal: XLAlignmentHorizontalValues.Center);
            SetCell(worksheet, $"I{row}", emission.GrossEmission, horizontal: XLAlignmentHorizontalValues.Center);
        }
        
        SetBorder(worksheet, $"A{row - report.Emissions.Count - 2}:I{row - report.Emissions.Count - 1}", XLBorderStyleValues.Medium);
        SetBorder(worksheet, $"A{row - report.Emissions.Count}:I{row - report.Emissions.Count}", XLBorderStyleValues.Medium);
        SetBorder(worksheet, $"A{row - report.Emissions.Count + 1}:I{row}", XLBorderStyleValues.Medium);
        row += 2;
        
        SetCell(worksheet, $"A{row}:I{row}", "Результаты расчетов", true);
        row++;
        
        SetCell(worksheet, $"A{row}", "Код ЗВ", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"B{row}:E{row}", "Наименование ЗВ", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"F{row}:G{row}", "г/с", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"H{row}:I{row}", "т/г", true, XLAlignmentHorizontalValues.Center);

        SetBorder(worksheet, $"A{row}", XLBorderStyleValues.Medium);
        SetBorder(worksheet, $"B{row}:I{row}", XLBorderStyleValues.Medium);
        SetBorder(worksheet, $"A{row + 1}:A{row + report.Emissions.Count}", XLBorderStyleValues.Medium);
        SetBorder(worksheet, $"B{row + 1}:I{row + report.Emissions.Count}", XLBorderStyleValues.Medium);

        foreach (var emission in report.Emissions)
        {
            row++;
            SetCell(worksheet, $"A{row}", emission.PollutantInfo.Code, horizontal: XLAlignmentHorizontalValues.Center);
            SetCell(worksheet, $"B{row}:E{row}", emission.PollutantInfo.Name, horizontal: XLAlignmentHorizontalValues.Center);
            SetCell(worksheet, $"F{row}:G{row}", emission.MaximumEmission, horizontal: XLAlignmentHorizontalValues.Center);
            SetCell(worksheet, $"H{row}:I{row}", emission.GrossEmission, horizontal: XLAlignmentHorizontalValues.Center);
        }

        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0; 
        return stream;
    }

    public MemoryStream CreateReservoirsEmissionsReport(ReservoirsEmissionsReport report)
    {
        var sheetName = $"ИЗА {report.PollutionSource}_{report.SelectionSource}";
        
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(sheetName);

        var row = 1;
        
        SetCell(worksheet, $"A{row}:H{row}", "Расчет выбросов загрязняющих веществ от резервуаров", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"A{row + 2}:H{row + 2}", $"Источник загрязнения № {report.PollutionSource}, резервуары для дизельного топлива", true);
        SetCell(worksheet, $"A{row + 3}:H{row + 3}", $"Источник выделения № {report.SelectionSource}", true);
        SetCell(worksheet, $"A{row + 4}:H{row + 4}", "Литература: Методические указания по определению выбросов загрязняющих веществ в атмосферу из резервуаров (утверждены приказом Госкомэкологии России от 08.04.1998 \u2116 199)");
        row += 6;
        
        SetCell(worksheet, $"A{row}:H{row}", "Исходные данные", true);
        SetCell(worksheet, $"A{row + 1}:D{row + 1}", "Конструкция резервуара");
        SetCell(worksheet, $"E{row + 1}:H{row + 1}", $"{report.ReservoirType.GetDescription()} ({report.ClimateZone.GetDescription()})");

        SetCell(worksheet, $"A{row + 2}:D{row + 2}", "Объем резервуара, м3");
        SetCell(worksheet, $"E{row + 2}:H{row + 2}", report.ReservoirVolume);

        SetCell(worksheet, $"A{row + 3}:D{row + 3}", "Количество резервуаров, шт.");
        SetCell(worksheet, $"E{row + 3}:H{row + 3}", report.ReservoirCount);

        SetCell(worksheet, $"A{row + 4}:D{row + 4}", "Время работы, ч/г");
        SetCell(worksheet, $"E{row + 4}:H{row + 4}", report.WorkHoursPerYear);
        
        SetBorder(worksheet, $"A{row + 1}:H{row + 4}");
        row += 6;
        
        SetCell(worksheet, $"A{row}:H{row}", "Расчет выброса", true);
        SetCell(worksheet, $"A{row + 1}:D{row + 1}", "Максимальные выбросы паров нефтепродуктов, М (г/с)", true);
        SetCell(worksheet, $"E{row + 1}:H{row + 1}", "М = (Сpmax \u2219Vсл) : 1200", true);
        SetCell(worksheet, $"A{row + 2}:D{row + 2}", "Максимальная концентрация паров нефтепродуктов в резервуаре Сpmax, (г/м3) (прил. 15)");
        SetCell(worksheet, $"E{row + 2}:H{row + 2}", report.VaporConcentration.MaxVaporConcentration);
        
        SetCell(worksheet, $"A{row + 3}:D{row + 3}", "Объем слитого нефтепродукта в резервуар,  Vсл (м3)");
        SetCell(worksheet, $"E{row + 3}:H{row + 3}", report.DrainedVolume);
        
        SetCell(worksheet, $"A{row + 4}:D{row + 4}", "Среднее время слива, с");
        SetCell(worksheet, $"E{row + 4}:H{row + 4}", report.AverageDrainTime);
        
        SetCell(worksheet, $"A{row + 5}:D{row + 5}", "Количество закачиваемого в резервуар нефтепродукта в осенне-зимний период Qоз, (м3)");
        SetCell(worksheet, $"E{row + 5}:H{row + 5}", report.AutumnWinterOilAmount);
        
        SetCell(worksheet, $"A{row + 6}:D{row + 6}", "Количество закачиваемого в резервуар нефтепродукта в весенне-летний период Qвл, (м3)");
        SetCell(worksheet, $"E{row + 6}:H{row + 6}", report.SpringSummerOilAmount);
        
        SetCell(worksheet, $"A{row + 7}:D{row + 7}", "Концентрация паров нефтепродуктов при заполнении резервуаров осенне-зимний период, Соз  (г/м3)");
        SetCell(worksheet, $"E{row + 7}:H{row + 7}", report.VaporConcentration.AutumnWinterVaporConcentration);
        
        SetCell(worksheet, $"A{row + 8}:D{row + 8}", "Концентрация паров нефтепродуктов при заполнении резервуаров весенне-летний период, Свл  (г/м3)");
        SetCell(worksheet, $"E{row + 8}:H{row + 8}", report.VaporConcentration.SpringSummerVaporConcentration);
        
        SetCell(worksheet, $"A{row + 9}:D{row + 9}", "Максимальные выбросы паров нефтепродуктов, М (г/с)", true);
        SetCell(worksheet, $"E{row + 9}:H{row + 9}", report.Result.MaxVaporEmission, true);
        
        SetBorder(worksheet, $"A{row + 1}:H{row + 9}");
        row += 11;
        
        SetCell(worksheet, $"A{row}:D{row}", "Валовый выброс, G (т/г)", true);
        SetCell(worksheet, $"E{row}:H{row}", "G = Gзак + Gпр", true);
        
        SetCell(worksheet, $"A{row + 1}:D{row + 1}", "Годовые выбросы при закачке, Gзак (т/г)");
        SetCell(worksheet, $"E{row + 1}:H{row + 1}", "Gзак = [(Ср+Сб)\u2219Qоз+(Ср+Сб)\u2219Qвл] \u2219 10-6");
        
        SetCell(worksheet, $"A{row + 2}:D{row + 2}", "Годовые выбросы при закачке, Gзак (т/г)");
        SetCell(worksheet, $"E{row + 2}:H{row + 2}", "Gзак = (Ср\u2219Qоз + Ср\u2219Qвл) \u2219 10-6", true);
        
        SetCell(worksheet, $"A{row + 3}:D{row + 3}", "Годовые выбросы при закачке, Gзак (т/г)");
        SetCell(worksheet, $"E{row + 3}:H{row + 3}", report.Result.AnnualInjectionEmissions);
        
        SetCell(worksheet, $"A{row + 4}:D{row + 4}", "Годовые выбросы при проливе, Gпр (т/г) для дизтоплив");
        SetCell(worksheet, $"E{row + 4}:H{row + 4}", "Gпр = 50 \u2219 (Qоз + Qвл) \u2219 10-6", true);
        
        SetCell(worksheet, $"A{row + 5}:D{row + 5}", "Годовые выбросы при проливе, Gпр (т/г) для дизтоплив");
        SetCell(worksheet, $"E{row + 5}:H{row + 5}", report.Result.AnnualIrrigationEmissions);
        
        SetCell(worksheet, $"A{row + 6}:D{row + 6}", "Валовый выброс, G (т/г)", true);
        SetCell(worksheet, $"E{row + 6}:H{row + 6}", report.Result.AnnualInjectionEmissions + report.Result.AnnualIrrigationEmissions, true);
        
        SetBorder(worksheet, $"A{row}:H{row + 6}");
        row += 8;
        
        SetCell(worksheet, $"A{row}:D{row}", "Максимальный выброс i-го загрязняющего вещества, Мi (г/с)", true);
        SetCell(worksheet, $"E{row}:H{row}", "Мi = М \u2219 Сi \u2219 10-2", true);
        
        SetCell(worksheet, $"A{row + 1}:D{row + 1}", "Валовые выбросы, Gi (т/г)", true);
        SetCell(worksheet, $"E{row + 1}:H{row + 1}", "Gi = G \u2219 Ci \u2219 10-2", true);
        
        SetCell(worksheet, $"A{row + 2}:D{row + 2 + report.Result.Emissions.Count - 1}", "Концентрация i-го загрязняющего вещества,    Сi % мас. (прил. 14)");
        row++;
        
        foreach (var emission in report.Result.Emissions)
        {
            row++;
            SetCell(worksheet, $"E{row}:G{row}", emission.PollutantInfo.ShortName);
            SetCell(worksheet, $"H{row}", emission.PollutantInfo.SpecificEmission, numberFormat: "0.##");
        }
        
        SetBorder(worksheet, $"A{row - report.Result.Emissions.Count - 1}:H{row}");
        row += 2;
        
        SetCell(worksheet, $"A{row}:H{row}", "Итого по источнику", true);
        SetCell(worksheet, $"A{row + 1}", "Код ЗВ", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"B{row + 1}:D{row + 1}", "Наименование ЗВ", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"E{row + 1}:F{row + 1}", "Максимальный выброс, г/с", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"G{row + 1}:H{row + 1}", "Валовый выброс т/г", true, XLAlignmentHorizontalValues.Center);
        row++;
        
        foreach (var emission in report.Result.Emissions)
        {
            row++;
            SetCell(worksheet, $"A{row}", emission.PollutantInfo.Code, horizontal: XLAlignmentHorizontalValues.Center);
            SetCell(worksheet, $"B{row}:D{row}", emission.PollutantInfo.Name, horizontal: XLAlignmentHorizontalValues.Center);
            SetCell(worksheet, $"E{row}:F{row}", emission.MaximumEmission, horizontal: XLAlignmentHorizontalValues.Center);
            SetCell(worksheet, $"G{row}:H{row}", emission.GrossEmission, horizontal: XLAlignmentHorizontalValues.Center);
        }
        
        SetBorder(worksheet, $"A{row - report.Result.Emissions.Count}:H{row}");
        
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0; 
        return stream;
    }

    public MemoryStream CreateDuringMetalMachiningEmissionsReport(DuringMetalMachiningEmissionsReport report)
    {
        var sheetName = $"ИЗА {report.PollutionSource}_{report.SelectionSource}";
        
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(sheetName);

        var row = 1;
        
        SetCell(worksheet, $"A{row}:I{row}", "Расчет выбросов загрязняющих веществ при механической обработке металлов", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"A{row + 2}:I{row + 2}", $"Источник загрязнения № {report.PollutionSource}, {report.MetalMachiningMachineType.GetDescription().ToLower()}", true);
        SetCell(worksheet, $"A{row + 3}:I{row + 3}", $"Источник выделения № {report.SelectionSource}", true);
        SetCell(worksheet, $"A{row + 4}:I{row + 7}", "Литература: Методика расчета выделений (выбросов) загрязняющих веществ в атмосферу при механической обработке металлов (на основе удельных показателей) (утверждена приказом Госкомэкологии от 14.04.1997 \u2116 158)");
        row += 9;
        
        SetCell(worksheet, $"A{row}:I{row}", "Исходные данные", true);
        SetCell(worksheet, $"A{row + 1}:C{row + 1}", "Вид оборудования:");
        SetCell(worksheet, $"D{row + 1}:I{row + 1}", report.MetalMachiningMachineType.GetDescription().ToLower());
        
        SetCell(worksheet, $"A{row + 2}:C{row + 2}", "Тип охлаждения:");
        SetCell(worksheet, $"D{row + 2}:I{row + 2}", "нет");
        
        SetCell(worksheet, $"A{row + 3}:C{row + 3}", "Вид обрабатываемого материала:");
        SetCell(worksheet, $"D{row + 3}:I{row + 3}", "сталь");
        
        SetCell(worksheet, $"A{row + 4}:C{row + 4}", "Годовой фонд времени работы оборудования, ч (Т)");
        SetCell(worksheet, $"D{row + 4}:I{row + 4}", report.WorkDaysPerYear);
        
        SetCell(worksheet, $"A{row + 5}:C{row + 5}", "Число оборудования данного типа (n)");
        SetCell(worksheet, $"D{row + 5}:I{row + 5}", report.MachiningMachineCount);
        
        SetCell(worksheet, $"A{row + 6}:C{row + 6}", "Число оборудования данного типа работающего одновременно (n)");
        SetCell(worksheet, $"D{row + 6}:I{row + 6}", report.SameMachiningMachineCount);
        
        SetBorder(worksheet, $"A{row + 1}:I{row + 6}");
        row += 8;
        
        SetCell(worksheet, $"A{row}:I{row}", "Расчет выброса", true);
        SetCell(worksheet, $"A{row + 1}:E{row + 1}", "Валовое значение мощности выделений и выбросов ЗВ для i-го ИЗА, т/г");
        SetCell(worksheet, $"F{row + 1}:I{row + 1}", "Miгв=0,2\u22193,6\u2219qi\u2219Т\u221910-3, т/г");
        
        SetCell(worksheet, $"A{row + 2}:E{row + 2}", "Максимально-разовый выброс ЗВ, г/с");
        SetCell(worksheet, $"F{row + 2}:I{row + 2}", "Mi в=0,2\u2219qi, г/с");
        
        SetCell(worksheet, $"A{row + 3}:E{row + 3}", "qi- удельные выделения пыли технологическим оборудованием (табл. П.2.1), г/с:");
        SetCell(worksheet, $"F{row + 3}:H{row + 3}", "пыль металлическая");
        SetCell(worksheet, $"I{row + 3}", report.Result.PollutantInfo.SpecificEmission);
        
        SetBorder(worksheet, $"A{row + 1}:I{row + 3}");
        row += 5;
        
        SetCell(worksheet, $"A{row}:I{row}", "Результат расчета", true);
        SetCell(worksheet, $"A{row + 1}", "Код ЗВ", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"B{row + 1}:E{row + 1}", "Наименование ЗВ", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"F{row + 1}:G{row + 1}", "Максимальный выброс, г/с", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"H{row + 1}:I{row + 1}", "Валовый выброс, т/г", true, XLAlignmentHorizontalValues.Center);
        
        SetCell(worksheet, $"A{row + 2}", report.Result.PollutantInfo.Code, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"B{row + 2}:E{row + 2}", report.Result.PollutantInfo.Name, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"F{row + 2}:G{row + 2}", report.Result.MaximumEmission, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"H{row + 2}:I{row + 2}", report.Result.GrossEmission, horizontal: XLAlignmentHorizontalValues.Center);
        
        SetBorder(worksheet, $"A{row + 1}:I{row + 2}");
        
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0; 
        return stream;
    }
    
    public MemoryStream CreateDuringWeldingOperationsEmissionsReport(DuringWeldingOperationsEmissionsReport report)
    {
        var sheetName = $"ИЗА {report.PollutionSource}_{report.SelectionSource}";
        
        var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(sheetName);

        var row = 1;
        
        SetCell(worksheet, $"A{row}:G{row}", "Расчет выбросов загрязняющих веществ при сварочных работах", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"A{row + 2}:G{row + 2}", $"Источник загрязнения № {report.PollutionSource}, Сварочный аппарат", true);
        SetCell(worksheet, $"A{row + 3}:G{row + 3}", $"Источник выделения № {report.SelectionSource}", true);
        SetCell(worksheet, $"A{row + 4}:G{row + 4}", "Литература: \"Методика расчета выделений (выбросов) загрязняющих веществ в атмосферу при сварочных работах (на основе удельных показателей)\", Госкомэкологии России от 14.04.1997 \u2116 158) ");
        row += 6;
        
        SetCell(worksheet, $"A{row}:G{row}", "Исходные данные", true);
        SetCell(worksheet, $"A{row + 1}:B{row + 1}", "Марка электродов");
        SetCell(worksheet, $"C{row + 1}:G{row + 1}", "ОК 46.00 (по аналогу МР-3)", true, XLAlignmentHorizontalValues.Center);
        
        SetCell(worksheet, $"A{row + 2}:B{row + 2}", "Расход сварочных электродов в год, кг");
        SetCell(worksheet, $"C{row + 2}:G{row + 2}", report.ElectrodesPerYear, true, XLAlignmentHorizontalValues.Center);
        
        SetCell(worksheet, $"A{row + 3}:B{row + 3}", "Расход сварочных электродов в год с учетом нормативного образования огарков, кг");
        SetCell(worksheet, $"C{row + 3}:G{row + 3}", "Вэ=G\u2219(100-н)\u221910-2", true, XLAlignmentHorizontalValues.Center);
        
        SetCell(worksheet, $"A{row + 4}:B{row + 4}", "Расход сварочных электродов в год с учетом нормативного образования огарков, кг");
        SetCell(worksheet, $"C{row + 4}:G{row + 4}", report.Result.NormElectrodesPerYear, true, XLAlignmentHorizontalValues.Center);
        
        SetCell(worksheet, $"A{row + 5}:B{row + 5}", "Расход сварочных материалов, кг/час (B)");
        UpdateCell(worksheet, $"A{row + 5}:B{row + 5}");
        SetCell(worksheet, $"C{row + 5}:G{row + 5}", report.Result.MaterialsConsumption, true, XLAlignmentHorizontalValues.Center);
        
        SetCell(worksheet, $"A{row + 6}:B{row + 6}", "Время работы сварочного оборудования, ч/год (T)");
        UpdateCell(worksheet, $"A{row + 6}:B{row + 6}");
        SetCell(worksheet, $"C{row + 6}:G{row + 6}", report.WorkDaysPerYear, true, XLAlignmentHorizontalValues.Center);
        
        SetCell(worksheet, $"A{row + 7}:B{row + 7}", "Наименование загрязняющего вещества");
        SetCell(worksheet, $"A{row + 8}:B{row + 8}", "Удельный показатель выделения i-го загрязняющего вещества на единицу массы расходуемых сырья и материалов, г/кг (КМ i)");
        UpdateCell(worksheet, $"A{row + 8}:B{row + 8}", 6);

        var column = 'B';
        foreach (var emission in report.Result.Emissions)
        {
            column = GetNextLetter(column);
            SetCell(worksheet, $"{column}{row + 7}", $"{emission.PollutantInfo.Code} {emission.PollutantInfo.Name}", true, XLAlignmentHorizontalValues.Center);
            SetCell(worksheet, $"{column}{row + 8}", emission.PollutantInfo.SpecificEmission,  horizontal: XLAlignmentHorizontalValues.Center);
        }

        if (report.Result.Emissions.Count < 5)
        {
            worksheet.Range($"{column}{row + 7}:G{row + 7}").Merge();
            worksheet.Range($"{column}{row + 8}:G{row + 8}").Merge();
        }

        SetCell(worksheet, $"A{row + 9}:B{row + 10}", "Эффективность местной установки очистки газов, в долях единицы (ɳ):");
        UpdateCell(worksheet, $"A{row + 9}:B{row + 10}", 4);
        SetCell(worksheet, $"C{row + 9}:D{row + 9}", "для твердых веществ:");
        SetCell(worksheet, $"C{row + 10}:D{row + 10}", "для газообразных веществ:");
        SetCell(worksheet, $"E{row + 9}:G{row + 9}", 0, horizontal: XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"E{row + 10}:G{row + 10}", 0, horizontal: XLAlignmentHorizontalValues.Center);
        
        SetBorder(worksheet, $"A{row + 1}:G{row + 10}");
        row += 12;
        
        SetCell(worksheet, $"A{row}:B{row}", "Наименование", true);
        SetCell(worksheet, $"C{row}:G{row}", "Расчетная формула", true);
        SetCell(worksheet, $"A{row + 1}:B{row + 1}", "Максимальный выброс, Ммi (г/с) рассчитывается по формуле:");
        SetCell(worksheet, $"C{row + 1}:G{row + 1}", "Ммi=B\u2219Кмi\u2219(1-ɳ)\u2219(1-ɳ1i)\u2219Кгр/3600", true);
        SetCell(worksheet, $"A{row + 2}:B{row + 2}", "Валовый выброс, МГ1Мi, (т/г) рассчитывается по формуле:");
        SetCell(worksheet, $"C{row + 2}:G{row + 2}", "МГ1Мi = 3,6\u2219MМi\u2219Т\u22190.001", true);
        SetBorder(worksheet, $"A{row}:G{row + 2}");
        row += 4;
        
        SetCell(worksheet, $"A{row}:G{row}", "Результаты расчетов", true);
        SetCell(worksheet, $"A{row + 1}", "Код ЗВ", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"B{row + 1}:E{row + 1}", "Наименование ЗВ", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"F{row + 1}", "Максимальный выброс, г/с", true, XLAlignmentHorizontalValues.Center);
        SetCell(worksheet, $"G{row + 1}", "Валовый выброс, т/г", true, XLAlignmentHorizontalValues.Center);
        row++;
        
        foreach (var emission in report.Result.Emissions)
        {
            row++;
            SetCell(worksheet, $"A{row}", emission.PollutantInfo.Code, horizontal: XLAlignmentHorizontalValues.Center);
            SetCell(worksheet, $"B{row}:E{row}", emission.PollutantInfo.Name, horizontal: XLAlignmentHorizontalValues.Center);
            SetCell(worksheet, $"F{row}", emission.MaximumEmission, horizontal: XLAlignmentHorizontalValues.Center);
            SetCell(worksheet, $"G{row}", emission.GrossEmission, horizontal: XLAlignmentHorizontalValues.Center);
        }
        SetBorder(worksheet, $"A{row - report.Result.Emissions.Count}:G{row}");
        
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0; 
        return stream;
    }

    private static void SetCell(IXLWorksheet worksheet, string address, object value, bool bold = false, XLAlignmentHorizontalValues horizontal = XLAlignmentHorizontalValues.Left, string numberFormat = "0.######")
    {
        if (address.Contains(':'))
        {
            worksheet.Range(address).Merge();
            address = address.Split(':')[0];
        }
        
        var cell = worksheet.Cell(address);

        cell.Value = value switch
        {
            string stringValue => stringValue,
            int intValue => intValue,
            float floatValue => floatValue,
            _ => cell.Value
        };

        cell.Style.NumberFormat.Format = numberFormat;
        cell.Style.Alignment.Horizontal = horizontal;
        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        cell.Style.Font.Bold = bold;
        cell.Style.Alignment.WrapText = true;
        cell.Style.Font.FontName = "Times New Roman";
        cell.Style.Font.FontSize = 12;
    }

    private static void UpdateCell(IXLWorksheet worksheet, string address, int boldLength = 3)
    {
        if (address.Contains(':'))
        {
            address = address.Split(':')[0];
        }
        
        var cell = worksheet.Cell(address);
        var richText = cell.CreateRichText();
        var length = richText.Text.Length;
        richText.Substring(length - boldLength, boldLength).SetBold();
    }

    private static void SetBorder(IXLWorksheet worksheet, string address, XLBorderStyleValues style = XLBorderStyleValues.Thin)
    {
        var range = worksheet.Range(address);
        range.Style.Border.OutsideBorder = style;
        range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
    }
    
    private static char GetNextLetter(char c)
    {
        return c switch
        {
            'z' => 'a',
            'Z' => 'A',
            _ => (char)(c + 1)
        };
    }
}