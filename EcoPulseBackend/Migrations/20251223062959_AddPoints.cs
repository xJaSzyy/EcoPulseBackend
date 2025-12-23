using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace EcoPulseBackend.Migrations
{
    /// <inheritdoc />
    public partial class AddPoints : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "EndLat",
                table: "VehicleFlowEmissionSources");

            migrationBuilder.DropColumn(
                name: "EndLon",
                table: "VehicleFlowEmissionSources");

            migrationBuilder.DropColumn(
                name: "StartLat",
                table: "VehicleFlowEmissionSources");

            migrationBuilder.DropColumn(
                name: "StartLon",
                table: "VehicleFlowEmissionSources");

            migrationBuilder.AddColumn<string>(
                name: "Points",
                table: "VehicleFlowEmissionSources",
                type: "text",
                nullable: false,
                defaultValue: "");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "Points",
                table: "VehicleFlowEmissionSources");

            migrationBuilder.AddColumn<double>(
                name: "EndLat",
                table: "VehicleFlowEmissionSources",
                type: "double precision",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.AddColumn<double>(
                name: "EndLon",
                table: "VehicleFlowEmissionSources",
                type: "double precision",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.AddColumn<double>(
                name: "StartLat",
                table: "VehicleFlowEmissionSources",
                type: "double precision",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.AddColumn<double>(
                name: "StartLon",
                table: "VehicleFlowEmissionSources",
                type: "double precision",
                nullable: false,
                defaultValue: 0.0);
        }
    }
}
