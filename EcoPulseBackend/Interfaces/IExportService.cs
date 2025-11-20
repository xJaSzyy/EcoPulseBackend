using EcoPulseBackend.Models;

namespace EcoPulseBackend.Interfaces;

public interface IExportService
{
    public MemoryStream CreateGasolineGeneratorEmissionsReport(GasolineGeneratorEmissionsReport report);

    public MemoryStream CreateReservoirsEmissionsReport(ReservoirsEmissionsReport report);

    public MemoryStream CreateDuringMetalMachiningEmissionsReport(DuringMetalMachiningEmissionsReport report);

    public MemoryStream CreateDuringWeldingOperationsEmissionsReport(DuringWeldingOperationsEmissionsReport report);
}