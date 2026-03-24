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
    private GetDocumentBaseContentQueryHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _documentRepository = Substitute.For<IDocumentRepository>();
        _handler = new GetDocumentBaseContentQueryHandler(_documentRepository);
    }

    [Test]
    public async Task Handle_DocumentWithSingleVersion_ShouldReturnBaseContent()
    {
        // Arrange
        var masterId = Guid.NewGuid();
        var document = new Document(masterId, "contract.pdf", "application/pdf", "Admin");
        var content = new byte[] { 37, 80, 68, 70 }; // %PDF magic bytes
        document.AddVersion(content, "Admin");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

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
        document.AddVersion(originalContent, "User");
        document.AddVersion(new byte[] { 4, 5, 6 }, "User");
        document.AddVersion(new byte[] { 7, 8, 9 }, "User");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

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
        document.AddVersion([1, 2, 3], "User");

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

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
        var v1 = document.AddVersion(v1Content, "User");
        document.AddVersion([2], "User");
        document.AddVersion([3], "User");
        // Aktywna jest v3, ale base content = v1
        v1.IsActive.Should().BeFalse();

        _documentRepository.GetByIdWithVersionsAsync(masterId, Arg.Any<CancellationToken>())
            .Returns(document);

        var query = new GetDocumentBaseContentQuery(masterId);

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.Value!.Content.Should().BeEquivalentTo(v1Content);
        result.Value.FileName.Should().Contain("_v1");
    }
}
