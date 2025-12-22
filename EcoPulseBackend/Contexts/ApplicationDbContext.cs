using EcoPulseBackend.Models.SingleEmissionSource;
using EcoPulseBackend.Models.TrafficLightQueueEmissionSource;
using EcoPulseBackend.Models.VehicleFlowEmissionSource;
using Microsoft.EntityFrameworkCore;

namespace EcoPulseBackend.Contexts;

public class ApplicationDbContext : DbContext
{
    public virtual DbSet<SingleEmissionSource> SingleEmissionSources { get; set; } = null!;
    public virtual DbSet<VehicleFlowEmissionSource> VehicleFlowEmissionSources { get; set; } = null!;
    public virtual DbSet<TrafficLightQueueEmissionSource> TrafficLightQueueEmissionSources { get; set; } = null!;
    
    public ApplicationDbContext() {  }
    
    public ApplicationDbContext(DbContextOptions options) : base(options)
    {
        Database.EnsureCreated();
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
            .OwnsOne(
                e => e.StartLocation,
                cb =>
                {
                    cb.Property(c => c.Lon).HasColumnName("StartLon");
                    cb.Property(c => c.Lat).HasColumnName("StartLat");
                });

        builder.Entity<VehicleFlowEmissionSource>()
            .OwnsOne(
                e => e.EndLocation,
                cb =>
                {
                    cb.Property(c => c.Lon).HasColumnName("EndLon");
                    cb.Property(c => c.Lat).HasColumnName("EndLat");
                });
        
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