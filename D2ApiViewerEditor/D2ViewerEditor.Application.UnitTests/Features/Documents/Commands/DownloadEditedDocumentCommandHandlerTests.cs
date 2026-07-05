using D2ViewerEditor.Application.Features.Documents.Commands.DownloadEditedDocument;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Domain.Models;
using FluentAssertions;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class DownloadEditedDocumentCommandHandlerTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private Mock<IDocumentRepository> _documentRepo = null!;
    private Mock<IHtmlToDocxConverter> _converter = null!;
    private Mock<IDocumentStorageService> _storage = null!;
    private DownloadEditedDocumentCommandHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _documentRepo = new Mock<IDocumentRepository>();
        _converter = new Mock<IHtmlToDocxConverter>();
        _storage = new Mock<IDocumentStorageService>();
        _handler = new DownloadEditedDocumentCommandHandler(
            _documentRepo.Object,
            _converter.Object,
            _storage.Object,
            NullLogger<DownloadEditedDocumentCommandHandler>.Instance);
    }

    private static Document NewDoc(string? metadata) =>
        new(Guid.NewGuid(), "doc.docx", DocxMime, "User", metadata);

    private DownloadEditedDocumentCommand BuildCmd(Guid masterId, string html = "<p>x</p>") =>
        new(masterId, html, "out.docx", null, null, null, null);

    [Test]
    public async Task Handle_DocumentMissing_ReturnsNotFound()
    {
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(It.IsAny<Guid>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync((Document?)null);

        var result = await _handler.Handle(BuildCmd(Guid.NewGuid()), CancellationToken.None);

        result.IsNotFound.Should().BeTrue();
        _converter.Verify(c => c.Convert(
            It.IsAny<string>(), It.IsAny<DocumentMetadata?>(),
            It.IsAny<HeaderFooterContent?>(), It.IsAny<HeaderFooterContent?>(),
            It.IsAny<PageMargins?>(), It.IsAny<PageSize?>(), It.IsAny<IReadOnlyList<SectionHeaderFooter>?>()), Times.Never);
    }

    [TestCase(null, TestName = "metadata missing entirely")]
    [TestCase("", TestName = "metadata empty")]
    [TestCase("{}", TestName = "metadata without userDownload")]
    [TestCase("{\"userDownload\":false}", TestName = "userDownload=false")]
    [TestCase("{\"userDownload\":null}", TestName = "userDownload=null")]
    [TestCase("{\"userDownload\":\"true\"}", TestName = "userDownload as string")]
    [TestCase("not-json", TestName = "metadata malformed")]
    public async Task Handle_UserDownloadNotTrue_ReturnsForbidden_WithoutInvokingConverter(string? metadata)
    {
        var doc = NewDoc(metadata);
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);

        var result = await _handler.Handle(BuildCmd(doc.Id), CancellationToken.None);

        result.IsSuccess.Should().BeFalse();
        result.IsNotFound.Should().BeFalse();
        result.Error.Should().StartWith(DownloadEditedDocumentCommandHandler.ForbiddenErrorPrefix);
        _converter.Verify(c => c.Convert(
            It.IsAny<string>(), It.IsAny<DocumentMetadata?>(),
            It.IsAny<HeaderFooterContent?>(), It.IsAny<HeaderFooterContent?>(),
            It.IsAny<PageMargins?>(), It.IsAny<PageSize?>(), It.IsAny<IReadOnlyList<SectionHeaderFooter>?>()), Times.Never);
    }

    [Test]
    public async Task Handle_UserDownloadTrue_ConvertsAndReturnsBytes()
    {
        var doc = NewDoc("{\"userDownload\":true,\"classification\":\"C2\"}");
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);

        var expected = new byte[] { 1, 2, 3, 4 };
        _converter.Setup(c => c.Convert(
                It.IsAny<string>(), It.IsAny<DocumentMetadata?>(),
                It.IsAny<HeaderFooterContent?>(), It.IsAny<HeaderFooterContent?>(),
                It.IsAny<PageMargins?>(), It.IsAny<PageSize?>(), It.IsAny<IReadOnlyList<SectionHeaderFooter>?>()))
            .Returns(expected);

        var result = await _handler.Handle(BuildCmd(doc.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.DocxBytes.Should().Equal(expected);
        result.Value.FileName.Should().Be("out.docx");
    }

    [Test]
    public async Task Handle_AllowedButEmptyHtml_ReturnsFailure()
    {
        var doc = NewDoc("{\"userDownload\":true}");
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);

        var result = await _handler.Handle(BuildCmd(doc.Id, html: "   "), CancellationToken.None);

        result.IsSuccess.Should().BeFalse();
        result.Error.Should().NotStartWith(DownloadEditedDocumentCommandHandler.ForbiddenErrorPrefix);
    }

    [Test]
    public async Task Handle_WithOriginalDocxVersion_UsesPassThrough_NotFromScratch()
    {
        var doc = NewDoc("{\"userDownload\":true}");
        doc.AddVersion(Guid.NewGuid(), "documents/v1", 123, "User"); // base/original package
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);

        var originalBytes = new byte[] { 10, 20, 30 };
        _storage.Setup(s => s.DownloadAsync("documents/v1", It.IsAny<CancellationToken>())).ReturnsAsync(originalBytes);

        var expected = new byte[] { 7, 7, 7 };
        _converter.Setup(c => c.ConvertPreservingPackage(
                It.IsAny<string>(), It.IsAny<Stream?>(), It.IsAny<DocumentMetadata?>(),
                It.IsAny<HeaderFooterContent?>(), It.IsAny<HeaderFooterContent?>(),
                It.IsAny<PageMargins?>(), It.IsAny<PageSize?>(), It.IsAny<IReadOnlyList<SectionHeaderFooter>?>()))
            .Returns(expected);

        var result = await _handler.Handle(BuildCmd(doc.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.DocxBytes.Should().Equal(expected);
        _storage.Verify(s => s.DownloadAsync("documents/v1", It.IsAny<CancellationToken>()), Times.Once);
        _converter.Verify(c => c.Convert(
            It.IsAny<string>(), It.IsAny<DocumentMetadata?>(), It.IsAny<HeaderFooterContent?>(),
            It.IsAny<HeaderFooterContent?>(), It.IsAny<PageMargins?>(), It.IsAny<PageSize?>(), It.IsAny<IReadOnlyList<SectionHeaderFooter>?>()), Times.Never);
    }

    [Test]
    public async Task Handle_SectionHeadersFooters_AreForwardedToConvert()
    {
        // Bez wersji bazowej → ścieżka Convert; lista sekcji musi dojść nietknięta (ADR-0025).
        var doc = NewDoc("{\"userDownload\":true}");
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);

        var sectionHf = new List<SectionHeaderFooter>
        {
            new() { SectionIndex = 1, Footer = new HeaderFooterContent { Html = "<p>Stopka sekcji 2</p>" } }
        };
        _converter.Setup(c => c.Convert(
                It.IsAny<string>(), It.IsAny<DocumentMetadata?>(),
                It.IsAny<HeaderFooterContent?>(), It.IsAny<HeaderFooterContent?>(),
                It.IsAny<PageMargins?>(), It.IsAny<PageSize?>(), It.IsAny<IReadOnlyList<SectionHeaderFooter>?>()))
            .Returns(new byte[] { 5 });

        var cmd = new DownloadEditedDocumentCommand(doc.Id, "<p>x</p>", "out.docx", null, null, null, null,
            PageSize: null, SectionHeadersFooters: sectionHf);

        var result = await _handler.Handle(cmd, CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        _converter.Verify(c => c.Convert(
            It.IsAny<string>(), It.IsAny<DocumentMetadata?>(),
            It.IsAny<HeaderFooterContent?>(), It.IsAny<HeaderFooterContent?>(),
            It.IsAny<PageMargins?>(), It.IsAny<PageSize?>(),
            It.Is<IReadOnlyList<SectionHeaderFooter>?>(s => ReferenceEquals(s, sectionHf))), Times.Once);
    }

    [Test]
    public async Task Handle_SectionHeadersFooters_AreForwardedToPassThroughConvert()
    {
        // Z wersją bazową → ścieżka ConvertPreservingPackage; ta sama gwarancja przekazania listy.
        var doc = NewDoc("{\"userDownload\":true}");
        doc.AddVersion(Guid.NewGuid(), "documents/v1", 123, "User");
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);
        _storage.Setup(s => s.DownloadAsync("documents/v1", It.IsAny<CancellationToken>()))
            .ReturnsAsync(new byte[] { 1 });

        var sectionHf = new List<SectionHeaderFooter>
        {
            new() { SectionIndex = 1, Header = new HeaderFooterContent { Html = "<p>Nagłówek sekcji 2</p>" } }
        };
        _converter.Setup(c => c.ConvertPreservingPackage(
                It.IsAny<string>(), It.IsAny<Stream?>(), It.IsAny<DocumentMetadata?>(),
                It.IsAny<HeaderFooterContent?>(), It.IsAny<HeaderFooterContent?>(),
                It.IsAny<PageMargins?>(), It.IsAny<PageSize?>(), It.IsAny<IReadOnlyList<SectionHeaderFooter>?>()))
            .Returns(new byte[] { 7 });

        var cmd = new DownloadEditedDocumentCommand(doc.Id, "<p>x</p>", "out.docx", null, null, null, null,
            PageSize: null, SectionHeadersFooters: sectionHf);

        var result = await _handler.Handle(cmd, CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        _converter.Verify(c => c.ConvertPreservingPackage(
            It.IsAny<string>(), It.IsAny<Stream?>(), It.IsAny<DocumentMetadata?>(),
            It.IsAny<HeaderFooterContent?>(), It.IsAny<HeaderFooterContent?>(),
            It.IsAny<PageMargins?>(), It.IsAny<PageSize?>(),
            It.Is<IReadOnlyList<SectionHeaderFooter>?>(s => ReferenceEquals(s, sectionHf))), Times.Once);
    }

    [Test]
    public async Task Handle_NoOriginalFileName_FallsBackToDocumentNameWithDocxSuffix()
    {
        var doc = NewDoc("{\"userDownload\":true}");
        _documentRepo.Setup(r => r.GetByIdWithVersionsAsync(doc.Id, It.IsAny<CancellationToken>())).ReturnsAsync(doc);
        _converter.Setup(c => c.Convert(
                It.IsAny<string>(), It.IsAny<DocumentMetadata?>(),
                It.IsAny<HeaderFooterContent?>(), It.IsAny<HeaderFooterContent?>(),
                It.IsAny<PageMargins?>(), It.IsAny<PageSize?>(), It.IsAny<IReadOnlyList<SectionHeaderFooter>?>()))
            .Returns(new byte[] { 9 });

        var cmd = new DownloadEditedDocumentCommand(doc.Id, "<p>x</p>", null, null, null, null, null);

        var result = await _handler.Handle(cmd, CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.FileName.Should().Contain("doc.docx");
    }
}
