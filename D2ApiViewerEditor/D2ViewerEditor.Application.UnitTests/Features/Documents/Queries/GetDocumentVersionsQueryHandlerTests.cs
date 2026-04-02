using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentVersions;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Queries;

[TestFixture]
public class GetDocumentVersionsQueryHandlerTests
{
    private IDocumentRepository _documentRepository;
    private GetDocumentVersionsQueryHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _documentRepository = Substitute.For<IDocumentRepository>();
        _handler = new GetDocumentVersionsQueryHandler(_documentRepository);
    }

    [Test]
    public async Task Handle_DocumentWithMultipleVersions_ShouldReturnAllVersionsSortedByVersionNumber()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "changelog.md", "text/markdown", "Developer");
        
        var v1 = document.AddVersion("documents/test/v1", 10, "Developer");
        Thread.Sleep(10); // Ensure different timestamps
        var v2 = document.AddVersion("documents/test/v2", 20, "Developer");
        Thread.Sleep(10);
        var v3 = document.AddVersion("documents/test/v3", 30, "Developer");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

        var query = new GetDocumentVersionsQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value.Should().NotBeNull();
        result.Value.Should().HaveCount(3);
        
        // Sortowanie od najnowszej (3, 2, 1)
        result.Value![0].VersionNumber.Should().Be(3);
        result.Value[1].VersionNumber.Should().Be(2);
        result.Value[2].VersionNumber.Should().Be(1);

        // Sprawdź że tylko ostatnia jest aktywna
        result.Value[0].IsActive.Should().BeTrue();
        result.Value[1].IsActive.Should().BeFalse();
        result.Value[2].IsActive.Should().BeFalse();
    }

    [Test]
    public async Task Handle_DocumentNotFound_ShouldReturnFailure()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns((Document?)null);

        var query = new GetDocumentVersionsQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.IsNotFound.Should().BeTrue();
    }

    [Test]
    public async Task Handle_VersionsShouldNotIncludeContent()
    {
        // Arrange - DTO powinno zwracać tylko metadane, bez contentu
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "large_file.bin", "application/octet-stream", "User");
        
        var largeContent = new byte[10 * 1024 * 1024]; // 10 MB
        document.AddVersion("documents/test/v1", largeContent.Length, "User");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

        var query = new GetDocumentVersionsQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value.Should().HaveCount(1);
        result.Value![0].SizeInBytes.Should().Be(10 * 1024 * 1024);
        // DTO DocumentVersionDto nie zawiera pola Content
    }

    [Test]
    public async Task Handle_SingleVersion_ShouldReturnCorrectMetadata()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "readme.txt", "text/plain", "Admin");
        var content = new byte[] { 72, 101, 108, 108, 111 }; // "Hello"
        var version = document.AddVersion("documents/test/v1", content.Length, "Admin");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

        var query = new GetDocumentVersionsQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value.Should().HaveCount(1);
        
        var versionDto = result.Value![0];
        versionDto.VersionId.Should().Be(version.Id);
        versionDto.VersionNumber.Should().Be(1);
        versionDto.CreatedBy.Should().Be("Admin");
        versionDto.IsActive.Should().BeTrue();
        versionDto.SizeInBytes.Should().Be(5);
        versionDto.CreatedAt.Should().BeCloseTo(DateTime.UtcNow, TimeSpan.FromSeconds(5));
    }

    [Test]
    public async Task Handle_AfterRestore_ShouldShowCorrectActiveVersion()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "contract.pdf", "application/pdf", "Lawyer");
        
        var v1 = document.AddVersion("documents/test/v1", 10, "Lawyer");
        var v2 = document.AddVersion("documents/test/v2", 20, "Lawyer");
        var v3 = document.AddVersion("documents/test/v3", 30, "Lawyer");
        
        // Restore do wersji 1
        document.RestoreVersion(v1.Id);

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

        var query = new GetDocumentVersionsQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value.Should().HaveCount(3);
        
        // Wersja 1 powinna być aktywna mimo że nie jest najnowsza
        var activeVersion = result.Value!.Single(v => v.IsActive);
        activeVersion.VersionNumber.Should().Be(1);
        activeVersion.VersionId.Should().Be(v1.Id);
    }

    [Test]
    public async Task Handle_RepositoryException_ShouldReturnFailure()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        _documentRepository.When(x => x.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>()))
            .Do(x => throw new Exception("Network timeout"));

        var query = new GetDocumentVersionsQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("Network timeout");
    }

    [Test]
    public async Task Handle_EmptyVersionsList_ShouldReturnEmptyList()
    {
        // Arrange - teoretyczny edge case
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "empty.bin", "application/octet-stream", "System");
        // Brak wersji

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

        var query = new GetDocumentVersionsQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value.Should().BeEmpty();
    }
}
