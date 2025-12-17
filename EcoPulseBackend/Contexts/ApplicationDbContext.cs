using EcoPulseBackend.Models;
using Microsoft.EntityFrameworkCore;

namespace EcoPulseBackend.Contexts;

public class ApplicationDbContext : DbContext
{
    public virtual DbSet<EmissionSource> EmissionSources { get; set; } = null!;
    
    public ApplicationDbContext() {  }
    
    public ApplicationDbContext(DbContextOptions options) : base(options)
    {
        Database.EnsureCreated();
    }
}