using D2ViewerEditor.Domain.Entities;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Domain.UnitTests.Entities;

[TestFixture]
public class DocumentTests
{
    private static readonly string SampleStoragePath = "documents/test/v1";

    // ── Constructor ────────────────────────────────────────────────────────────

    [Test]
    public void Constructor_WithValidParameters_ShouldCreateDocument()
    {
        var id = Guid.NewGuid();
        var doc = new Document(id, "report.pdf", "application/pdf", "Admin");

        doc.Id.Should().Be(id);
        doc.Name.Should().Be("report.pdf");
        doc.MimeType.Should().Be("application/pdf");
        doc.CreatedBy.Should().Be("Admin");
        doc.IsDeleted.Should().BeFalse();
        doc.Versions.Should().BeEmpty();
        doc.CreatedAt.Should().BeCloseTo(DateTime.UtcNow, TimeSpan.FromSeconds(5));
    }

    [Test]
    public void Constructor_WithEmptyName_ShouldThrowArgumentException()
    {
        var act = () => new Document(Guid.NewGuid(), "", "application/pdf", "User");

        act.Should().Throw<ArgumentException>()
            .WithMessage("*Nazwa dokumentu*");
    }

    [Test]
    public void Constructor_WithWhitespaceName_ShouldThrowArgumentException()
    {
        var act = () => new Document(Guid.NewGuid(), "   ", "application/pdf", "User");

        act.Should().Throw<ArgumentException>();
    }

    [Test]
    public void Constructor_WithEmptyMimeType_ShouldThrowArgumentException()
    {
        var act = () => new Document(Guid.NewGuid(), "file.pdf", "", "User");

        act.Should().Throw<ArgumentException>()
            .WithMessage("*MimeType*");
    }

    // ── AddVersion ─────────────────────────────────────────────────────────────

    [Test]
    public void AddVersion_FirstVersion_ShouldBeNumberedOneAndActive()
    {
        var doc = new Document(Guid.NewGuid(), "doc.pdf", "application/pdf", "User");

        var version = doc.AddVersion(SampleStoragePath, 500, "User");

        version.VersionNumber.Should().Be(1);
        version.IsActive.Should().BeTrue();
        version.StoragePath.Should().Be(SampleStoragePath);
        doc.Versions.Should().HaveCount(1);
    }

    [Test]
    public void AddVersion_SecondVersion_ShouldDeactivatePreviousAndBeActive()
    {
        var doc = new Document(Guid.NewGuid(), "doc.pdf", "application/pdf", "User");
        var v1 = doc.AddVersion("documents/test/v1", 100, "User");

        var v2 = doc.AddVersion("documents/test/v2", 200, "User");

        v1.IsActive.Should().BeFalse();
        v2.IsActive.Should().BeTrue();
        v2.VersionNumber.Should().Be(2);
    }

    [Test]
    public void AddVersion_MultipleVersions_ShouldIncrementVersionNumbers()
    {
        var doc = new Document(Guid.NewGuid(), "doc.pdf", "application/pdf", "User");

        var versions = Enumerable.Range(1, 5)
            .Select(i => doc.AddVersion($"documents/test/v{i}", i * 10, "User"))
            .ToList();

        versions.Select(v => v.VersionNumber).Should().BeEquivalentTo([1, 2, 3, 4, 5]);
        versions.Last().IsActive.Should().BeTrue();
        versions.Take(4).Should().AllSatisfy(v => v.IsActive.Should().BeFalse());
    }

    [Test]
    public void AddVersion_WithEmptyStoragePath_ShouldThrowArgumentException()
    {
        var doc = new Document(Guid.NewGuid(), "doc.pdf", "application/pdf", "User");

        var act = () => doc.AddVersion("", 10, "User");

        act.Should().Throw<ArgumentException>()
            .WithMessage("*storage*");
    }

    [Test]
    public void AddVersion_ShouldAssignUniqueIds()
    {
        var doc = new Document(Guid.NewGuid(), "doc.pdf", "application/pdf", "User");
        doc.AddVersion("documents/test/v1", 10, "User");
        doc.AddVersion("documents/test/v2", 20, "User");

        var ids = doc.Versions.Select(v => v.Id).ToList();
        ids.Should().OnlyHaveUniqueItems();
        ids.Should().NotContain(Guid.Empty);
    }

    [Test]
    public void AddVersion_ShouldLinkVersionToDocumentId()
    {
        var docId = Guid.NewGuid();
        var doc = new Document(docId, "doc.pdf", "application/pdf", "User");

        var version = doc.AddVersion(SampleStoragePath, 500, "User");

        version.DocumentId.Should().Be(docId);
    }

    // ── GetActiveVersion ───────────────────────────────────────────────────────

    [Test]
    public void GetActiveVersion_WithNoVersions_ShouldReturnNull()
    {
        var doc = new Document(Guid.NewGuid(), "doc.pdf", "application/pdf", "User");

        doc.GetActiveVersion().Should().BeNull();
    }

    [Test]
    public void GetActiveVersion_WithMultipleVersions_ShouldReturnLatest()
    {
        var doc = new Document(Guid.NewGuid(), "doc.pdf", "application/pdf", "User");
        doc.AddVersion("documents/test/v1", 10, "User");
        doc.AddVersion("documents/test/v2", 20, "User");
        var v3 = doc.AddVersion("documents/test/v3", 30, "User");

        doc.GetActiveVersion()!.Id.Should().Be(v3.Id);
    }

    // ── RestoreVersion ─────────────────────────────────────────────────────────

    [Test]
    public void RestoreVersion_ValidId_ShouldActivateAndDeactivateOthers()
    {
        var doc = new Document(Guid.NewGuid(), "doc.pdf", "application/pdf", "User");
        var v1 = doc.AddVersion("documents/test/v1", 10, "User");
        var v2 = doc.AddVersion("documents/test/v2", 20, "User");
        var v3 = doc.AddVersion("documents/test/v3", 30, "User");

        doc.RestoreVersion(v1.Id);

        v1.IsActive.Should().BeTrue();
        v2.IsActive.Should().BeFalse();
        v3.IsActive.Should().BeFalse();
    }

    [Test]
    public void RestoreVersion_CurrentActiveVersion_ShouldRemainActive()
    {
        var doc = new Document(Guid.NewGuid(), "doc.pdf", "application/pdf", "User");
        doc.AddVersion("documents/test/v1", 10, "User");
        var v2 = doc.AddVersion("documents/test/v2", 20, "User");

        doc.RestoreVersion(v2.Id);

        v2.IsActive.Should().BeTrue();
    }

    [Test]
    public void RestoreVersion_NonExistentId_ShouldThrowInvalidOperationException()
    {
        var doc = new Document(Guid.NewGuid(), "doc.pdf", "application/pdf", "User");
        doc.AddVersion("documents/test/v1", 10, "User");

        var act = () => doc.RestoreVersion(Guid.NewGuid());

        act.Should().Throw<InvalidOperationException>()
            .WithMessage("*nie istnieje*");
    }

    [Test]
    public void RestoreVersion_ShouldPreserveAllVersionsInHistory()
    {
        var doc = new Document(Guid.NewGuid(), "doc.pdf", "application/pdf", "User");
        var v1 = doc.AddVersion("documents/test/v1", 10, "User");
        doc.AddVersion("documents/test/v2", 20, "User");
        doc.AddVersion("documents/test/v3", 30, "User");

        doc.RestoreVersion(v1.Id);

        doc.Versions.Should().HaveCount(3);
    }

    // ── ChangeName ─────────────────────────────────────────────────────────────

    [Test]
    public void ChangeName_WithValidName_ShouldUpdateName()
    {
        var doc = new Document(Guid.NewGuid(), "old.pdf", "application/pdf", "User");

        doc.ChangeName("new-name.pdf");

        doc.Name.Should().Be("new-name.pdf");
    }

    [Test]
    public void ChangeName_WithEmptyName_ShouldThrowArgumentException()
    {
        var doc = new Document(Guid.NewGuid(), "old.pdf", "application/pdf", "User");

        var act = () => doc.ChangeName("");

        act.Should().Throw<ArgumentException>()
            .WithMessage("*Nazwa*");
    }

    // ── Delete ─────────────────────────────────────────────────────────────────

    [Test]
    public void Delete_ShouldSetIsDeletedToTrue()
    {
        var doc = new Document(Guid.NewGuid(), "doc.pdf", "application/pdf", "User");

        doc.Delete();

        doc.IsDeleted.Should().BeTrue();
    }

    [Test]
    public void Delete_CalledTwice_ShouldRemainDeleted()
    {
        var doc = new Document(Guid.NewGuid(), "doc.pdf", "application/pdf", "User");

        doc.Delete();
        doc.Delete();

        doc.IsDeleted.Should().BeTrue();
    }
}
