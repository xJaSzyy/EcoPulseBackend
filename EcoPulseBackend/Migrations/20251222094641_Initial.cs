using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

#nullable disable

namespace EcoPulseBackend.Migrations
{
    /// <inheritdoc />
    public partial class Initial : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "SingleEmissionSources",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    Lon = table.Column<double>(type: "double precision", nullable: false),
                    Lat = table.Column<double>(type: "double precision", nullable: false),
                    EjectedTemp = table.Column<float>(type: "real", nullable: false),
                    AvgExitSpeed = table.Column<float>(type: "real", nullable: false),
                    HeightSource = table.Column<float>(type: "real", nullable: false),
                    DiameterSource = table.Column<float>(type: "real", nullable: false),
                    TempStratificationRatio = table.Column<int>(type: "integer", nullable: false),
                    SedimentationRateRatio = table.Column<int>(type: "integer", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_SingleEmissionSources", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "TrafficLightQueueEmissionSources",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    Lon = table.Column<double>(type: "double precision", nullable: false),
                    Lat = table.Column<double>(type: "double precision", nullable: false),
                    VehicleType = table.Column<int>(type: "integer", nullable: false),
                    VehiclesCount = table.Column<int>(type: "integer", nullable: false),
                    TrafficLightCycles = table.Column<int>(type: "integer", nullable: false),
                    TrafficLightStopTime = table.Column<float>(type: "real", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_TrafficLightQueueEmissionSources", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "VehicleFlowEmissionSources",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    StartLon = table.Column<double>(type: "double precision", nullable: false),
                    StartLat = table.Column<double>(type: "double precision", nullable: false),
                    EndLon = table.Column<double>(type: "double precision", nullable: false),
                    EndLat = table.Column<double>(type: "double precision", nullable: false),
                    VehicleType = table.Column<int>(type: "integer", nullable: false),
                    MaxTrafficIntensity = table.Column<float>(type: "real", nullable: false),
                    AverageSpeed = table.Column<float>(type: "real", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_VehicleFlowEmissionSources", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "SingleEmissionSources");

            migrationBuilder.DropTable(
                name: "TrafficLightQueueEmissionSources");

            migrationBuilder.DropTable(
                name: "VehicleFlowEmissionSources");
        }
    }
}
