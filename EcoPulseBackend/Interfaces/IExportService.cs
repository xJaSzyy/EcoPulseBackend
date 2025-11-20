using EcoPulseBackend.Models;

namespace EcoPulseBackend.Interfaces;

public interface IExportService
{
    public MemoryStream CreateGasolineGeneratorEmissionsReport(GasolineGeneratorEmissionsReport report, string fileName);

    public MemoryStream CreateReservoirsEmissionsReport(ReservoirsEmissionsReport report, string fileName);

    public MemoryStream CreateDuringMetalMachiningEmissionsReport(DuringMetalMachiningEmissionsReport report, string fileName);

    public MemoryStream CreateDuringWeldingOperationsEmissionsReport(DuringWeldingOperationsEmissionsReport report, string fileName);
}