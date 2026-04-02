using D2ViewerEditor.Application.Features.Documents.Queries.GetDocument;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Queries;

[TestFixture]
public class GetDocumentQueryHandlerTests
{
    private IDocumentRepository _documentRepository;
    private IDocumentStorageService _storageService;
    private GetDocumentQueryHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _documentRepository = Substitute.For<IDocumentRepository>();
        _storageService = Substitute.For<IDocumentStorageService>();
        _handler = new GetDocumentQueryHandler(_documentRepository, _storageService);
    }

    [Test]
    public async Task Handle_ExistingDocument_ShouldReturnActiveVersion()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "invoice.pdf", "application/pdf", "Accountant");
        var content = new byte[] { 10, 20, 30, 40 };
        var version = document.AddVersion("documents/test/v1", content.Length, "Accountant");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);
        _storageService.DownloadAsync(version.StoragePath, Arg.Any<CancellationToken>())
            .Returns(content);

        var query = new GetDocumentQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value.Should().NotBeNull();
        result.Value!.MasterId.Should().Be(masterId);
        result.Value.Name.Should().Be("invoice.pdf");
        result.Value.MimeType.Should().Be("application/pdf");
        result.Value.Content.Should().BeEquivalentTo(content);
        result.Value.VersionNumber.Should().Be(1);
    }

    [Test]
    public async Task Handle_DocumentNotFound_ShouldReturnFailure()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns((Document?)null);

        var query = new GetDocumentQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.IsNotFound.Should().BeTrue();
    }

    [Test]
    public async Task Handle_DocumentWithMultipleVersions_ShouldReturnActiveOne()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Analyst");
        
        document.AddVersion("documents/test/v1", 20, "Analyst");
        document.AddVersion("documents/test/v2", 20, "Analyst");
        var activeContent = new byte[] { 5, 6, 7 };
        var activeVersion = document.AddVersion("documents/test/v3", activeContent.Length, "Analyst"); // Wersja 3 - aktywna

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);
        _storageService.DownloadAsync(activeVersion.StoragePath, Arg.Any<CancellationToken>())
            .Returns(activeContent);

        var query = new GetDocumentQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Content.Should().BeEquivalentTo(activeContent);
        result.Value.VersionNumber.Should().Be(3);
    }

    [Test]
    public async Task Handle_DocumentWithNoActiveVersion_ShouldReturnFailure()
    {
        // Arrange - teoretyczny edge case (nie powinien wystąpić w rzeczywistości)
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "corrupted.bin", "application/octet-stream", "System");
        // Nie dodajemy żadnej wersji lub wszystkie są nieaktywne

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

        var query = new GetDocumentQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("nie ma aktywnej wersji");
    }

    [Test]
    public async Task Handle_RepositoryException_ShouldReturnFailure()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        _documentRepository.When(x => x.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>()))
            .Do(x => throw new Exception("Connection lost"));

        var query = new GetDocumentQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("Connection lost");
    }
}
