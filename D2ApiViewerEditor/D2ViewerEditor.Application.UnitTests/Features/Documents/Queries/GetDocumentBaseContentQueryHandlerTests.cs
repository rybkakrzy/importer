using D2ViewerEditor.Application.Common.Security;
using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentBaseContent;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Queries;

[TestFixture]
public class GetDocumentBaseContentQueryHandlerTests
{
    private IDocumentRepository _documentRepository;
    private IDocumentStorageService _storageService;
    private IDocumentAccessGuard _accessGuard;
    private GetDocumentBaseContentQueryHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _documentRepository = Substitute.For<IDocumentRepository>();
        _storageService = Substitute.For<IDocumentStorageService>();
        _accessGuard = Substitute.For<IDocumentAccessGuard>();
        _accessGuard.IsViewAllowed(Arg.Any<string?>()).Returns(true);
        _handler = new GetDocumentBaseContentQueryHandler(_documentRepository, _storageService, _accessGuard);
    }

    [Test]
    public async Task Handle_DocumentWithSingleVersion_ShouldReturnBaseContent()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "contract.pdf", "application/pdf", "Admin");
        var content = new byte[] { 37, 80, 68, 70 }; // %PDF magic bytes
        var version = document.AddVersion("documents/test/v1", content.Length, "Admin");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);
        _storageService.DownloadAsync(version.StoragePath, Arg.Any<CancellationToken>())
            .Returns(content);

        var query = new GetDocumentBaseContentQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Content.Should().BeEquivalentTo(content);
        result.Value.MimeType.Should().Be("application/pdf");
        result.Value.FileName.Should().Be("contract.pdf_v1.pdf");
    }

    [Test]
    public async Task Handle_DocumentWithMultipleVersions_ShouldReturnFirstVersion()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "report.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "User");

        var originalContent = new byte[] { 1, 2, 3 };
        var baseVersion = document.AddVersion("documents/test/v1", originalContent.Length, "User");
        document.AddVersion("documents/test/v2", 30, "User");
        document.AddVersion("documents/test/v3", 30, "User");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);
        _storageService.DownloadAsync(baseVersion.StoragePath, Arg.Any<CancellationToken>())
            .Returns(originalContent);

        var query = new GetDocumentBaseContentQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Content.Should().BeEquivalentTo(originalContent);
        result.Value.FileName.Should().Contain("_v1");
    }

    [Test]
    public async Task Handle_DocumentNotFound_ShouldReturnNotFound()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns((Document?)null);

        var query = new GetDocumentBaseContentQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.IsNotFound.Should().BeTrue();
    }

    [Test]
    public async Task Handle_DocumentWithNoVersions_ShouldReturnFailure()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "empty.pdf", "application/pdf", "User");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

        var query = new GetDocumentBaseContentQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("nie ma żadnych wersji");
    }

    [TestCase("application/vnd.openxmlformats-officedocument.wordprocessingml.document", "_v1.docx")]
    [TestCase("application/msword", "_v1.doc")]
    [TestCase("application/pdf", "_v1.pdf")]
    [TestCase("application/octet-stream", "_v1")]
    public async Task Handle_ShouldReturnCorrectExtensionForMimeType(string mimeType, string expectedSuffix)
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "file", mimeType, "User");
        var version = document.AddVersion("documents/test/v1", 30, "User");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);
        _storageService.DownloadAsync(version.StoragePath, Arg.Any<CancellationToken>())
            .Returns(new byte[] { 1, 2, 3 });

        var query = new GetDocumentBaseContentQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.FileName.Should().EndWith(expectedSuffix);
        result.Value.MimeType.Should().Be(mimeType);
    }

    [Test]
    public async Task Handle_ShouldNotReturnActiveVersionUnlessItIsFirst()
    {
        // Arrange — v3 jest aktywna (po restore back do v1 — ale v1 jest pierwsza), 
        // endpoint zawsze zwraca wersję #1 niezależnie od aktywności
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "doc.pdf", "application/pdf", "User");
        var v1Content = new byte[] { 1 };
        var v1 = document.AddVersion("documents/test/v1", v1Content.Length, "User");
        document.AddVersion("documents/test/v2", 20, "User");
        document.AddVersion("documents/test/v3", 30, "User");
        // Aktywna jest v3, ale base content = v1
        v1.IsActive.Should().BeFalse();

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);
        _storageService.DownloadAsync(v1.StoragePath, Arg.Any<CancellationToken>())
            .Returns(v1Content);

        var query = new GetDocumentBaseContentQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.Value!.Content.Should().BeEquivalentTo(v1Content);
        result.Value.FileName.Should().Contain("_v1");
    }
}
