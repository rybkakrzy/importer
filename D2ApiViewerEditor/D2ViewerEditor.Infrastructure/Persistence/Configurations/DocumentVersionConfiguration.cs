using D2ViewerEditor.Domain.Entities;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;

namespace D2ViewerEditor.Infrastructure.Persistence.Configurations;

/// <summary>
/// Konfiguracja Entity Framework dla encji DocumentVersion
/// </summary>
public class DocumentVersionConfiguration : IEntityTypeConfiguration<DocumentVersion>
{
    public void Configure(EntityTypeBuilder<DocumentVersion> builder)
    {
        builder.ToTable("document_versions");

        builder.HasKey(v => v.Id);

        builder.Property(v => v.Id)
            .HasColumnName("id")
            .ValueGeneratedNever()   // Guid generowany w domenie (AddVersion). Bez tego EF stosuje
            .IsRequired();           // konwencję ValueGeneratedOnAdd i nową wersję z ustawionym kluczem
                                     // dodaną do śledzonego dokumentu traktuje jako Modified (UPDATE→0 wierszy).

        builder.Property(v => v.DocumentId)
            .HasColumnName("document_id")
            .IsRequired();

        builder.Property(v => v.StoragePath)
            .HasColumnName("storage_path")
            .HasMaxLength(500)
            .IsRequired();

        builder.Property(v => v.SizeInBytes)
            .HasColumnName("size_in_bytes")
            .IsRequired();

        builder.Property(v => v.VersionNumber)
            .HasColumnName("version_number")
            .IsRequired();

        builder.Property(v => v.CreatedAt)
            .HasColumnName("created_at")
            .IsRequired();

        builder.Property(v => v.CreatedBy)
            .HasColumnName("created_by")
            .HasMaxLength(255)
            .IsRequired();

        builder.Property(v => v.IsActive)
            .HasColumnName("is_active")
            .IsRequired();

        builder.Property(v => v.ModifiedAt)
            .HasColumnName("modified_at");

        // Indeksy
        builder.HasIndex(v => v.DocumentId);
        builder.HasIndex(v => new { v.DocumentId, v.IsActive });
        builder.HasIndex(v => v.CreatedAt);
    }
}
