using D2ViewerEditor.Domain.Entities;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Domain.UnitTests.Entities;

[TestFixture]
public class DocumentVersionTests
{
    [Test]
    public void Constructor_WithValidParameters_ShouldCreateVersion()
    {
        var id = Guid.NewGuid();
        var documentId = Guid.NewGuid();
        var storagePath = "documents/test/v1";

        var version = new DocumentVersion(id, documentId, storagePath, 300, 1, "Author");

        version.Id.Should().Be(id);
        version.DocumentId.Should().Be(documentId);
        version.StoragePath.Should().Be(storagePath);
        version.SizeInBytes.Should().Be(300);
        version.VersionNumber.Should().Be(1);
        version.CreatedBy.Should().Be("Author");
        version.IsActive.Should().BeTrue();
        version.CreatedAt.Should().BeCloseTo(DateTime.UtcNow, TimeSpan.FromSeconds(5));
    }

    [Test]
    public void Constructor_WithEmptyStoragePath_ShouldThrowArgumentException()
    {
        var act = () => new DocumentVersion(Guid.NewGuid(), Guid.NewGuid(), "", 10, 1, "User");

        act.Should().Throw<ArgumentException>()
            .WithMessage("*storage*");
    }

    [Test]
    public void Constructor_WithNullStoragePath_ShouldThrowArgumentException()
    {
        var act = () => new DocumentVersion(Guid.NewGuid(), Guid.NewGuid(), null!, 10, 1, "User");

        act.Should().Throw<ArgumentException>();
    }

    [Test]
    public void SizeInBytes_ShouldReturnConfiguredValue()
    {
        var version = new DocumentVersion(Guid.NewGuid(), Guid.NewGuid(), "documents/test/v1", 1024, 1, "User");

        version.SizeInBytes.Should().Be(1024);
    }

    [Test]
    public void SizeInBytes_WithSmallValue_ShouldReturnCorrectly()
    {
        var version = new DocumentVersion(Guid.NewGuid(), Guid.NewGuid(), "documents/test/v1", 1, 1, "User");

        version.SizeInBytes.Should().Be(1);
    }

    [Test]
    public void NewVersion_IsAlwaysActive()
    {
        var version = new DocumentVersion(Guid.NewGuid(), Guid.NewGuid(), "documents/test/v1", 10, 1, "User");

        version.IsActive.Should().BeTrue();
    }
}
