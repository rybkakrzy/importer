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
        var content = new byte[] { 10, 20, 30 };

        var version = new DocumentVersion(id, documentId, content, 1, "Author");

        version.Id.Should().Be(id);
        version.DocumentId.Should().Be(documentId);
        version.Content.Should().BeEquivalentTo(content);
        version.VersionNumber.Should().Be(1);
        version.CreatedBy.Should().Be("Author");
        version.IsActive.Should().BeTrue();
        version.CreatedAt.Should().BeCloseTo(DateTime.UtcNow, TimeSpan.FromSeconds(5));
    }

    [Test]
    public void Constructor_WithEmptyContent_ShouldThrowArgumentException()
    {
        var act = () => new DocumentVersion(Guid.NewGuid(), Guid.NewGuid(), [], 1, "User");

        act.Should().Throw<ArgumentException>()
            .WithMessage("*Zawartość*");
    }

    [Test]
    public void Constructor_WithNullContent_ShouldThrowArgumentException()
    {
        var act = () => new DocumentVersion(Guid.NewGuid(), Guid.NewGuid(), null!, 1, "User");

        act.Should().Throw<ArgumentException>();
    }

    [Test]
    public void SizeInBytes_ShouldReturnContentLength()
    {
        var content = new byte[1024];
        var version = new DocumentVersion(Guid.NewGuid(), Guid.NewGuid(), content, 1, "User");

        version.SizeInBytes.Should().Be(1024);
    }

    [Test]
    public void SizeInBytes_WithSingleByte_ShouldReturnOne()
    {
        var version = new DocumentVersion(Guid.NewGuid(), Guid.NewGuid(), [42], 1, "User");

        version.SizeInBytes.Should().Be(1);
    }

    [Test]
    public void NewVersion_IsAlwaysActive()
    {
        var version = new DocumentVersion(Guid.NewGuid(), Guid.NewGuid(), [1], 1, "User");

        version.IsActive.Should().BeTrue();
    }
}
