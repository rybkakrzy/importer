using D2ViewerEditor.Domain.Entities;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;

namespace D2ViewerEditor.Infrastructure.Persistence.Configurations;

/// <summary>
/// Konfiguracja Entity Framework dla encji Document
/// </summary>
public class DocumentConfiguration : IEntityTypeConfiguration<Document>
{
    public void Configure(EntityTypeBuilder<Document> builder)
    {
        builder.ToTable("documents");

        builder.HasKey(d => d.Id);

        builder.Property(d => d.Id)
            .HasColumnName("id")
            .IsRequired();

        builder.Property(d => d.Name)
            .HasColumnName("name")
            .HasMaxLength(500)
            .IsRequired();

        builder.Property(d => d.MimeType)
            .HasColumnName("mime_type")
            .HasMaxLength(100)
            .IsRequired();

        builder.Property(d => d.CreatedAt)
            .HasColumnName("created_at")
            .IsRequired();

        builder.Property(d => d.CreatedBy)
            .HasColumnName("created_by")
            .HasMaxLength(255)
            .IsRequired();

        builder.Property(d => d.IsDeleted)
            .HasColumnName("is_deleted")
            .IsRequired()
            .HasDefaultValue(false);

        // Relacja 1:N z DocumentVersion — backing field _versions
        builder.HasMany(d => d.Versions)
            .WithOne(v => v.Document!)
            .HasForeignKey(v => v.DocumentId)
            .OnDelete(DeleteBehavior.Cascade);

        builder.Navigation(d => d.Versions)
            .HasField("_versions")
            .UsePropertyAccessMode(PropertyAccessMode.Field);

        // Indeksy
        builder.HasIndex(d => d.CreatedAt);
        builder.HasIndex(d => d.IsDeleted);
    }
}
