using EcoPulseBackend.Models;
using EcoPulseBackend.Models.DuringMetalMachining;
using EcoPulseBackend.Models.DuringWeldingOperations;
using EcoPulseBackend.Models.GasolineGenerator;
using EcoPulseBackend.Models.Reservoirs;

namespace EcoPulseBackend.Interfaces;

public interface IExportService
{
    public MemoryStream CreateGasolineGeneratorEmissionsReport(GasolineGeneratorEmissionsReport report);

    public MemoryStream CreateReservoirsEmissionsReport(ReservoirsEmissionsReport report);

    public MemoryStream CreateDuringMetalMachiningEmissionsReport(DuringMetalMachiningEmissionsReport report);

    public MemoryStream CreateDuringWeldingOperationsEmissionsReport(DuringWeldingOperationsEmissionsReport report);
}