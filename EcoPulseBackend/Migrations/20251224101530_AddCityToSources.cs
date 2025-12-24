using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace EcoPulseBackend.Migrations
{
    /// <inheritdoc />
    public partial class AddCityToSources : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "CityId",
                table: "VehicleFlowEmissionSources",
                type: "integer",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "CityId",
                table: "TrafficLightQueueEmissionSources",
                type: "integer",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "CityId",
                table: "SingleEmissionSources",
                type: "integer",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<double>(
                name: "Lat",
                table: "Cities",
                type: "double precision",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.AddColumn<double>(
                name: "Lon",
                table: "Cities",
                type: "double precision",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.CreateIndex(
                name: "IX_VehicleFlowEmissionSources_CityId",
                table: "VehicleFlowEmissionSources",
                column: "CityId");

            migrationBuilder.CreateIndex(
                name: "IX_TrafficLightQueueEmissionSources_CityId",
                table: "TrafficLightQueueEmissionSources",
                column: "CityId");

            migrationBuilder.CreateIndex(
                name: "IX_SingleEmissionSources_CityId",
                table: "SingleEmissionSources",
                column: "CityId");

            migrationBuilder.AddForeignKey(
                name: "FK_SingleEmissionSources_Cities_CityId",
                table: "SingleEmissionSources",
                column: "CityId",
                principalTable: "Cities",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_TrafficLightQueueEmissionSources_Cities_CityId",
                table: "TrafficLightQueueEmissionSources",
                column: "CityId",
                principalTable: "Cities",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);

            migrationBuilder.AddForeignKey(
                name: "FK_VehicleFlowEmissionSources_Cities_CityId",
                table: "VehicleFlowEmissionSources",
                column: "CityId",
                principalTable: "Cities",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_SingleEmissionSources_Cities_CityId",
                table: "SingleEmissionSources");

            migrationBuilder.DropForeignKey(
                name: "FK_TrafficLightQueueEmissionSources_Cities_CityId",
                table: "TrafficLightQueueEmissionSources");

            migrationBuilder.DropForeignKey(
                name: "FK_VehicleFlowEmissionSources_Cities_CityId",
                table: "VehicleFlowEmissionSources");

            migrationBuilder.DropIndex(
                name: "IX_VehicleFlowEmissionSources_CityId",
                table: "VehicleFlowEmissionSources");

            migrationBuilder.DropIndex(
                name: "IX_TrafficLightQueueEmissionSources_CityId",
                table: "TrafficLightQueueEmissionSources");

            migrationBuilder.DropIndex(
                name: "IX_SingleEmissionSources_CityId",
                table: "SingleEmissionSources");

            migrationBuilder.DropColumn(
                name: "CityId",
                table: "VehicleFlowEmissionSources");

            migrationBuilder.DropColumn(
                name: "CityId",
                table: "TrafficLightQueueEmissionSources");

            migrationBuilder.DropColumn(
                name: "CityId",
                table: "SingleEmissionSources");

            migrationBuilder.DropColumn(
                name: "Lat",
                table: "Cities");

            migrationBuilder.DropColumn(
                name: "Lon",
                table: "Cities");
        }
    }
}
