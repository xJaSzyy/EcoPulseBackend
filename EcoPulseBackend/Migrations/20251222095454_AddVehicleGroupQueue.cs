using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

#nullable disable

namespace EcoPulseBackend.Migrations
{
    /// <inheritdoc />
    public partial class AddVehicleGroupQueue : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "VehicleType",
                table: "TrafficLightQueueEmissionSources");

            migrationBuilder.DropColumn(
                name: "VehiclesCount",
                table: "TrafficLightQueueEmissionSources");

            migrationBuilder.CreateTable(
                name: "VehicleGroupQueues",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    VehicleType = table.Column<int>(type: "integer", nullable: false),
                    VehiclesCount = table.Column<int>(type: "integer", nullable: false),
                    TrafficLightQueueEmissionSourceId = table.Column<int>(type: "integer", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_VehicleGroupQueues", x => x.Id);
                    table.ForeignKey(
                        name: "FK_VehicleGroupQueues_TrafficLightQueueEmissionSources_Traffic~",
                        column: x => x.TrafficLightQueueEmissionSourceId,
                        principalTable: "TrafficLightQueueEmissionSources",
                        principalColumn: "Id");
                });

            migrationBuilder.CreateIndex(
                name: "IX_VehicleGroupQueues_TrafficLightQueueEmissionSourceId",
                table: "VehicleGroupQueues",
                column: "TrafficLightQueueEmissionSourceId");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "VehicleGroupQueues");

            migrationBuilder.AddColumn<int>(
                name: "VehicleType",
                table: "TrafficLightQueueEmissionSources",
                type: "integer",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "VehiclesCount",
                table: "TrafficLightQueueEmissionSources",
                type: "integer",
                nullable: false,
                defaultValue: 0);
        }
    }
}
