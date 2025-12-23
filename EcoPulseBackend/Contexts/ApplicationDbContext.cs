using System.Text.Json;
using EcoPulseBackend.Models;
using EcoPulseBackend.Models.SingleEmissionSource;
using EcoPulseBackend.Models.TrafficLightQueue;
using EcoPulseBackend.Models.TrafficLightQueueEmissionSource;
using EcoPulseBackend.Models.VehicleFlowEmissionSource;
using Microsoft.EntityFrameworkCore;

namespace EcoPulseBackend.Contexts;

public class ApplicationDbContext : DbContext
{
    public virtual DbSet<SingleEmissionSource> SingleEmissionSources { get; set; } = null!;
    public virtual DbSet<VehicleFlowEmissionSource> VehicleFlowEmissionSources { get; set; } = null!;
    public virtual DbSet<TrafficLightQueueEmissionSource> TrafficLightQueueEmissionSources { get; set; } = null!;
    public virtual DbSet<VehicleGroupQueue> VehicleGroupQueues { get; set; } = null!;
    
    public ApplicationDbContext() {  }
    
    public ApplicationDbContext(DbContextOptions options) : base(options)
    {
        //Database.EnsureCreated();
    }

    protected override void OnModelCreating(ModelBuilder builder)
    {
        base.OnModelCreating(builder);

        builder.Entity<SingleEmissionSource>()
            .OwnsOne(
                e => e.Location,
                cb =>
                {
                    cb.Property(c => c.Lon).HasColumnName("Lon");
                    cb.Property(c => c.Lat).HasColumnName("Lat");
                });
        
        builder.Entity<VehicleFlowEmissionSource>()
            .Property(e => e.Points)
            .HasConversion(
                v => JsonSerializer.Serialize(v, (JsonSerializerOptions?)null),
                v => JsonSerializer.Deserialize<List<Coordinates>>(v, (JsonSerializerOptions?)null)!);

        
        builder.Entity<TrafficLightQueueEmissionSource>()
            .OwnsOne(
                e => e.Location,
                cb =>
                {
                    cb.Property(c => c.Lon).HasColumnName("Lon");
                    cb.Property(c => c.Lat).HasColumnName("Lat");
                });
    }
}