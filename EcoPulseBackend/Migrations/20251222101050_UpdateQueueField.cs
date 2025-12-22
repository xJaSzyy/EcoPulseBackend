using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace EcoPulseBackend.Migrations
{
    /// <inheritdoc />
    public partial class UpdateQueueField : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_VehicleGroupQueues_TrafficLightQueueEmissionSources_Traffic~",
                table: "VehicleGroupQueues");

            migrationBuilder.AlterColumn<int>(
                name: "TrafficLightQueueEmissionSourceId",
                table: "VehicleGroupQueues",
                type: "integer",
                nullable: false,
                defaultValue: 0,
                oldClrType: typeof(int),
                oldType: "integer",
                oldNullable: true);

            migrationBuilder.AddForeignKey(
                name: "FK_VehicleGroupQueues_TrafficLightQueueEmissionSources_Traffic~",
                table: "VehicleGroupQueues",
                column: "TrafficLightQueueEmissionSourceId",
                principalTable: "TrafficLightQueueEmissionSources",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_VehicleGroupQueues_TrafficLightQueueEmissionSources_Traffic~",
                table: "VehicleGroupQueues");

            migrationBuilder.AlterColumn<int>(
                name: "TrafficLightQueueEmissionSourceId",
                table: "VehicleGroupQueues",
                type: "integer",
                nullable: true,
                oldClrType: typeof(int),
                oldType: "integer");

            migrationBuilder.AddForeignKey(
                name: "FK_VehicleGroupQueues_TrafficLightQueueEmissionSources_Traffic~",
                table: "VehicleGroupQueues",
                column: "TrafficLightQueueEmissionSourceId",
                principalTable: "TrafficLightQueueEmissionSources",
                principalColumn: "Id");
        }
    }
}
