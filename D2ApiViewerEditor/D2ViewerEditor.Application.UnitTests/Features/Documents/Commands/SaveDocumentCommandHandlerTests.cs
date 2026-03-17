using D2ViewerEditor.Application.Features.Documents.Commands.SaveDocument;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Domain.Models;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class SaveDocumentCommandHandlerTests
{
    private IHtmlToDocxConverter _converter;
    private SaveDocumentCommandHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _converter = Substitute.For<IHtmlToDocxConverter>();
        _handler = new SaveDocumentCommandHandler(_converter);
    }

    [Test]
    public async Task Handle_ValidCommand_ShouldReturnDocxBytes()
    {
        // Arrange
        var expectedBytes = new byte[] { 1, 2, 3, 4 };
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata>(),
            Arg.Any<HeaderFooterContent>(), Arg.Any<HeaderFooterContent>(), Arg.Any<PageMargins>())
            .Returns(expectedBytes);

        var command = new SaveDocumentCommand(
            Html: "<p>Test</p>",
            OriginalFileName: "test.docx",
            Metadata: null,
            Header: null,
            Footer: null
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.DocxBytes.Should().Equal(expectedBytes);
        result.Value.FileName.Should().Be("test.docx");
    }

    [Test]
    public async Task Handle_WithoutOriginalFileName_ShouldUseDefaultName()
    {
        // Arrange
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata>(),
            Arg.Any<HeaderFooterContent>(), Arg.Any<HeaderFooterContent>(), Arg.Any<PageMargins>())
            .Returns(new byte[] { 1 });

        var command = new SaveDocumentCommand(
            Html: "<p>Test</p>",
            OriginalFileName: null,
            Metadata: null,
            Header: null,
            Footer: null
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.FileName.Should().Be("dokument.docx");
    }

    [Test]
    public async Task Handle_WithEmptyOriginalFileName_ShouldUseDefaultName()
    {
        // Arrange
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata>(),
            Arg.Any<HeaderFooterContent>(), Arg.Any<HeaderFooterContent>(), Arg.Any<PageMargins>())
            .Returns(new byte[] { 1 });

        var command = new SaveDocumentCommand(
            Html: "<p>Test</p>",
            OriginalFileName: "",
            Metadata: null,
            Header: null,
            Footer: null
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.FileName.Should().Be("dokument.docx");
    }

    [Test]
    public async Task Handle_FileNameWithoutDocxExtension_ShouldAppendDocxExtension()
    {
        // Arrange
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata>(),
            Arg.Any<HeaderFooterContent>(), Arg.Any<HeaderFooterContent>(), Arg.Any<PageMargins>())
            .Returns(new byte[] { 1 });

        var command = new SaveDocumentCommand(
            Html: "<p>Test</p>",
            OriginalFileName: "dokument",
            Metadata: null,
            Header: null,
            Footer: null
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.FileName.Should().Be("dokument.docx");
    }

    [Test]
    public async Task Handle_FileNameAlreadyHasDocxExtension_ShouldNotDoubleExtension()
    {
        // Arrange
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata>(),
            Arg.Any<HeaderFooterContent>(), Arg.Any<HeaderFooterContent>(), Arg.Any<PageMargins>())
            .Returns(new byte[] { 1 });

        var command = new SaveDocumentCommand(
            Html: "<p>Test</p>",
            OriginalFileName: "raport.docx",
            Metadata: null,
            Header: null,
            Footer: null
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.FileName.Should().Be("raport.docx");
    }

    [Test]
    public async Task Handle_WithMetadataAndMargins_ShouldPassThemToConverter()
    {
        // Arrange
        var metadata = new DocumentMetadata { Title = "Tytuł" };
        var margins = new PageMargins { Top = 3.0, Bottom = 3.0, Left = 2.0, Right = 2.0 };
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata>(),
            Arg.Any<HeaderFooterContent>(), Arg.Any<HeaderFooterContent>(), Arg.Any<PageMargins>())
            .Returns(new byte[] { 1, 2 });

        var command = new SaveDocumentCommand(
            Html: "<p>Test</p>",
            OriginalFileName: "test.docx",
            Metadata: metadata,
            Header: null,
            Footer: null,
            Margins: margins
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        _converter.Received(1).Convert(
            Arg.Is("<p>Test</p>"),
            Arg.Is(metadata),
            Arg.Is<HeaderFooterContent?>(h => h == null),
            Arg.Is<HeaderFooterContent?>(f => f == null),
            Arg.Is(margins));
    }
}
