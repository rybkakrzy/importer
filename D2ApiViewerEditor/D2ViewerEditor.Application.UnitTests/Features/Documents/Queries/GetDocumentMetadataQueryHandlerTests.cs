using D2ViewerEditor.Application.Features.Documents.Queries.GetDocumentMetadata;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using Moq;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Queries;

[TestFixture]
public class GetDocumentMetadataQueryHandlerTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private Mock<IDocumentRepository> _documentRepo = null!;
    private GetDocumentMetadataQueryHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _documentRepo = new Mock<IDocumentRepository>();
        _handler = new GetDocumentMetadataQueryHandler(_documentRepo.Object);
    }

    private async Task<DocumentMetadataDto> HandleWithMetadataAsync(string? metadata)
    {
        var document = new Document(Guid.NewGuid(), "doc.docx", DocxMime, "User", metadata);
        _documentRepo.Setup(r => r.GetByIdAsync(document.Id, It.IsAny<CancellationToken>()))
            .ReturnsAsync(document);

        var result = await _handler.Handle(new GetDocumentMetadataQuery(document.Id), CancellationToken.None);
        result.IsSuccess.Should().BeTrue();
        return result.Value!;
    }

    [Test]
    public async Task Handle_DocumentMissing_ReturnsNotFound()
    {
        _documentRepo.Setup(r => r.GetByIdAsync(It.IsAny<Guid>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((Document?)null);

        var result = await _handler.Handle(new GetDocumentMetadataQuery(Guid.NewGuid()), CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
    }

    [Test]
    public async Task Handle_ClassificationPresent_IsReturned()
    {
        var dto = await HandleWithMetadataAsync("{\"returnUrl\":\"https://x/cb\",\"classification\":\"C2\"}");
        dto.Classification.Should().Be("C2");
    }

    [Test]
    public async Task Handle_UnknownClassificationValue_IsPassedThroughUnchanged()
    {
        // Handler does not enforce the C1..C4 dictionary — unknown values flow through.
        var dto = await HandleWithMetadataAsync("{\"classification\":\"TOP-SECRET\"}");
        dto.Classification.Should().Be("TOP-SECRET");
    }

    [Test]
    public async Task Handle_ClassificationFieldAbsent_IsNull()
    {
        var dto = await HandleWithMetadataAsync("{\"returnUrl\":\"https://x/cb\"}");
        dto.Classification.Should().BeNull();
    }

    [Test]
    public async Task Handle_ClassificationExplicitNull_IsNull()
    {
        var dto = await HandleWithMetadataAsync("{\"classification\":null}");
        dto.Classification.Should().BeNull();
    }

    [Test]
    public async Task Handle_ClassificationEmptyOrWhitespace_IsReturnedAsIs_AndDoesNotThrow()
    {
        var empty = await HandleWithMetadataAsync("{\"classification\":\"\"}");
        empty.Classification.Should().BeEmpty();

        var whitespace = await HandleWithMetadataAsync("{\"classification\":\"   \"}");
        whitespace.Classification.Should().Be("   ");
    }

    [Test]
    public async Task Handle_NoMetadata_ReturnsNullClassification()
    {
        (await HandleWithMetadataAsync(null)).Classification.Should().BeNull();
        (await HandleWithMetadataAsync("   ")).Classification.Should().BeNull();
    }

    [Test]
    public async Task Handle_MalformedJson_DoesNotThrow_ReturnsNullFields()
    {
        var dto = await HandleWithMetadataAsync("{ this is not valid json ");
        dto.Classification.Should().BeNull();
        dto.ReturnUrl.Should().BeNull();
        dto.UserDownload.Should().BeFalse("malformed metadata never grants download");
    }

    [Test]
    public async Task Handle_NoUserDownloadField_ReturnsFalse()
    {
        (await HandleWithMetadataAsync(null)).UserDownload.Should().BeFalse();
        (await HandleWithMetadataAsync("{}")).UserDownload.Should().BeFalse();
        (await HandleWithMetadataAsync("{\"classification\":\"C2\"}")).UserDownload.Should().BeFalse();
    }

    [TestCase("{\"userDownload\":false}")]
    [TestCase("{\"userDownload\":null}")]
    public async Task Handle_UserDownloadNotTrue_ReturnsFalse(string metadata)
    {
        (await HandleWithMetadataAsync(metadata)).UserDownload.Should().BeFalse();
    }

    [Test]
    public async Task Handle_UserDownloadTrue_ReturnsTrue()
    {
        var dto = await HandleWithMetadataAsync("{\"userDownload\":true,\"classification\":\"C2\"}");
        dto.UserDownload.Should().BeTrue();
        dto.Classification.Should().Be("C2");
    }
}
