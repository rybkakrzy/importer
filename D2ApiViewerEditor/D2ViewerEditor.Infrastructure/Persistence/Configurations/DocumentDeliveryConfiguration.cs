using D2ViewerEditor.Domain.Entities;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;

namespace D2ViewerEditor.Infrastructure.Persistence.Configurations;

/// <summary>
/// Konfiguracja Entity Framework dla encji DocumentDelivery (tabela document_deliveries).
/// </summary>
public class DocumentDeliveryConfiguration : IEntityTypeConfiguration<DocumentDelivery>
{
    public void Configure(EntityTypeBuilder<DocumentDelivery> builder)
    {
        builder.ToTable("document_deliveries");

        builder.HasKey(d => d.Id);

        builder.Property(d => d.Id)
            .HasColumnName("id")
            .ValueGeneratedNever()
            .IsRequired();

        builder.Property(d => d.DocumentId)
            .HasColumnName("document_id")
            .IsRequired();

        builder.Property(d => d.SourceVersionId)
            .HasColumnName("source_version_id")
            .IsRequired();

        builder.Property(d => d.SnapshotObjectName)
            .HasColumnName("snapshot_object_name")
            .IsRequired();

        builder.Property(d => d.SnapshotSizeBytes)
            .HasColumnName("snapshot_size_bytes")
            .IsRequired();

        builder.Property(d => d.SnapshotSha256)
            .HasColumnName("snapshot_sha256")
            .IsRequired();

        builder.Property(d => d.RecipientUrl)
            .HasColumnName("recipient_url")
            .IsRequired();

        builder.Property(d => d.Status)
            .HasColumnName("status")
            .HasConversion<string>()
            .HasMaxLength(32)
            .IsRequired()
            .HasDefaultValue(DeliveryStatus.Pending);

        builder.Property(d => d.AttemptCount)
            .HasColumnName("attempt_count")
            .IsRequired();

        builder.Property(d => d.CreatedAt)
            .HasColumnName("created_at")
            .IsRequired();

        builder.Property(d => d.UpdatedAt)
            .HasColumnName("updated_at")
            .IsRequired();

        builder.Property(d => d.FirstAttemptAt)
            .HasColumnName("first_attempt_at");

        builder.Property(d => d.LastAttemptAt)
            .HasColumnName("last_attempt_at");

        builder.Property(d => d.NextAttemptAt)
            .HasColumnName("next_attempt_at")
            .IsRequired();

        builder.Property(d => d.DeadlineAt)
            .HasColumnName("deadline_at")
            .IsRequired();

        builder.Property(d => d.LockedUntil)
            .HasColumnName("locked_until");

        builder.Property(d => d.LockedBy)
            .HasColumnName("locked_by")
            .HasMaxLength(128);

        builder.Property(d => d.LastError)
            .HasColumnName("last_error");

        builder.Property(d => d.CorrelationId)
            .HasColumnName("correlation_id")
            .IsRequired();

        builder.Property(d => d.CreatedBy)
            .HasColumnName("created_by")
            .HasMaxLength(255)
            .IsRequired();

        builder.HasIndex(d => d.NextAttemptAt);
        builder.HasIndex(d => d.Status);
        builder.HasIndex(d => d.DocumentId);
    }
}
