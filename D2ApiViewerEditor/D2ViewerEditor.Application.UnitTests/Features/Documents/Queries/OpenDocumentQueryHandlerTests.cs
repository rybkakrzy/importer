using D2ViewerEditor.Application.Features.Documents.Queries.OpenDocument;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Domain.Models;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Queries;

[TestFixture]
public class OpenDocumentQueryHandlerTests
{
    private IDocxToHtmlConverter _converter;
    private OpenDocumentQueryHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _converter = Substitute.For<IDocxToHtmlConverter>();
        _handler = new OpenDocumentQueryHandler(_converter);
    }

    [Test]
    public async Task Handle_ValidStream_ShouldReturnDocumentContent()
    {
        // Arrange
        var expectedContent = new DocumentContent
        {
            Html = "<p>Test</p>",
            Metadata = new DocumentMetadata { Title = "Dokument" }
        };
        _converter.Convert(Arg.Any<Stream>()).Returns(expectedContent);

        using var stream = new MemoryStream(new byte[] { 1, 2, 3 });
        var query = new OpenDocumentQuery(stream, "test.docx");

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Html.Should().Be("<p>Test</p>");
        result.Value.Metadata.Title.Should().Be("Dokument");
    }

    [Test]
    public async Task Handle_WhenMetadataTitleIsNull_ShouldUseFileName()
    {
        // Arrange
        var expectedContent = new DocumentContent
        {
            Html = "<p>Test</p>",
            Metadata = new DocumentMetadata { Title = null }
        };
        _converter.Convert(Arg.Any<Stream>()).Returns(expectedContent);

        using var stream = new MemoryStream(new byte[] { 1, 2, 3 });
        var query = new OpenDocumentQuery(stream, "moj_dokument.docx");

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Metadata.Title.Should().Be("moj_dokument");
    }

    [Test]
    public async Task Handle_WhenMetadataTitleIsSet_ShouldPreserveTitle()
    {
        // Arrange
        var expectedContent = new DocumentContent
        {
            Html = "<p>Test</p>",
            Metadata = new DocumentMetadata { Title = "Istniejący tytuł" }
        };
        _converter.Convert(Arg.Any<Stream>()).Returns(expectedContent);

        using var stream = new MemoryStream(new byte[] { 1, 2, 3 });
        var query = new OpenDocumentQuery(stream, "inna_nazwa.docx");

        // Act
        var result = await _handler.Handle(query, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.Metadata.Title.Should().Be("Istniejący tytuł");
    }

    [Test]
    public async Task Handle_ShouldCallConverterWithStream()
    {
        // Arrange
        var content = new DocumentContent { Html = "", Metadata = new DocumentMetadata() };
        _converter.Convert(Arg.Any<Stream>()).Returns(content);

        using var stream = new MemoryStream(new byte[] { 1, 2, 3 });
        var query = new OpenDocumentQuery(stream, "file.docx");

        // Act
        await _handler.Handle(query, CancellationToken.None);

        // Assert
        _converter.Received(1).Convert(Arg.Any<Stream>());
    }
}
