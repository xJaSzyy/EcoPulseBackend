using EcoPulseBackend.Models.SingleEmissionSource;
using EcoPulseBackend.Models.VehicleFlowEmissionSource;
using Microsoft.EntityFrameworkCore;

namespace EcoPulseBackend.Contexts;

public class ApplicationDbContext : DbContext
{
    public virtual DbSet<SingleEmissionSource> SingleEmissionSources { get; set; } = null!;
    public virtual DbSet<VehicleFlowEmissionSource> VehicleFlowEmissionSources { get; set; } = null!;
    
    public ApplicationDbContext() {  }
    
    public ApplicationDbContext(DbContextOptions options) : base(options)
    {
        Database.EnsureCreated();
    }
}