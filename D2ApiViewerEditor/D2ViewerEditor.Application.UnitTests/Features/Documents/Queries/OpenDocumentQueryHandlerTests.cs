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
    private IDocumentInputNormalizer _normalizer;
    private OpenDocumentQueryHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _converter = Substitute.For<IDocxToHtmlConverter>();
        _normalizer = Substitute.For<IDocumentInputNormalizer>();
        // Domyślnie: wejście jest poprawnym DOCX (pass-through) → handler woła konwerter.
        _normalizer.Normalize(Arg.Any<byte[]>(), Arg.Any<string?>())
            .Returns(DocumentInputResult.Success(new byte[] { 1, 2, 3 }));
        _handler = new OpenDocumentQueryHandler(_converter, _normalizer);
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

    [Test]
    public async Task Handle_WhenPasswordRequired_ReturnsSentinel_AndDoesNotConvert()
    {
        _normalizer.Normalize(Arg.Any<byte[]>(), Arg.Any<string?>())
            .Returns(DocumentInputResult.Failure(DocumentInputStatus.PasswordRequired));
        using var stream = new MemoryStream(new byte[] { 1, 2, 3 });

        var result = await _handler.Handle(new OpenDocumentQuery(stream, "tajne.docx"), CancellationToken.None);

        result.IsFailure.Should().BeTrue();
        result.Error.Should().Be(OpenDocumentQueryHandler.PasswordRequiredSentinel);
        _converter.DidNotReceive().Convert(Arg.Any<Stream>());
    }

    [Test]
    public async Task Handle_WhenWrongPassword_ReturnsSentinel()
    {
        _normalizer.Normalize(Arg.Any<byte[]>(), Arg.Any<string?>())
            .Returns(DocumentInputResult.Failure(DocumentInputStatus.WrongPassword));
        using var stream = new MemoryStream(new byte[] { 1, 2, 3 });

        var result = await _handler.Handle(new OpenDocumentQuery(stream, "tajne.docx", "złe"), CancellationToken.None);

        result.Error.Should().Be(OpenDocumentQueryHandler.WrongPasswordSentinel);
    }

    [Test]
    public async Task Handle_WhenBinaryLegacyDoc_ReturnsUnsupportedSentinel()
    {
        _normalizer.Normalize(Arg.Any<byte[]>(), Arg.Any<string?>())
            .Returns(DocumentInputResult.Failure(DocumentInputStatus.UnsupportedLegacyDoc));
        using var stream = new MemoryStream(new byte[] { 1, 2, 3 });

        var result = await _handler.Handle(new OpenDocumentQuery(stream, "stary.doc"), CancellationToken.None);

        result.Error.Should().Be(OpenDocumentQueryHandler.UnsupportedLegacyDocSentinel);
        _converter.DidNotReceive().Convert(Arg.Any<Stream>());
    }

    [Test]
    public async Task Handle_PassesPasswordToNormalizer()
    {
        using var stream = new MemoryStream(new byte[] { 1, 2, 3 });
        _converter.Convert(Arg.Any<Stream>())
            .Returns(new DocumentContent { Html = "", Metadata = new DocumentMetadata() });

        await _handler.Handle(new OpenDocumentQuery(stream, "f.docx", "sezam"), CancellationToken.None);

        _normalizer.Received(1).Normalize(Arg.Any<byte[]>(), "sezam");
    }
}
