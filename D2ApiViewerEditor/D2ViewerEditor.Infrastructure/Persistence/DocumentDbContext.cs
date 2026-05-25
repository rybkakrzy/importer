using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Infrastructure.Persistence.Configurations;
using Microsoft.EntityFrameworkCore;

namespace D2ViewerEditor.Infrastructure.Persistence;

/// <summary>
/// DbContext dla zarządzania dokumentami
/// </summary>
public class DocumentDbContext : DbContext
{
    public DocumentDbContext(DbContextOptions<DocumentDbContext> options) : base(options)
    {
    }

    public DbSet<Document> Documents => Set<Document>();
    public DbSet<DocumentVersion> DocumentVersions => Set<DocumentVersion>();
    public DbSet<DocumentDelivery> DocumentDeliveries => Set<DocumentDelivery>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        // Stosuj konfiguracje z osobnych klas
        modelBuilder.ApplyConfiguration(new DocumentConfiguration());
        modelBuilder.ApplyConfiguration(new DocumentVersionConfiguration());
        modelBuilder.ApplyConfiguration(new DocumentDeliveryConfiguration());
    }
}
