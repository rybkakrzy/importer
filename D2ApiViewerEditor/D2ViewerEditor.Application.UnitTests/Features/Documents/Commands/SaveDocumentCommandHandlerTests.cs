using D2ViewerEditor.Application.Features.Documents.Commands.SaveDocument;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Domain.Models;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class SaveDocumentCommandHandlerTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

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

    [Test]
    public async Task Handle_WithPageSizeAndSectionHeadersFooters_ShouldPassThemToConverter()
    {
        // Arrange — dokument wielosekcyjny: sekcja 1 ma własny nagłówek (ADR-0025).
        var pageSize = new PageSize { WidthCm = 21, HeightCm = 29.7, Orientation = "portrait" };
        var sectionHf = new List<SectionHeaderFooter>
        {
            new() { SectionIndex = 1, Header = new HeaderFooterContent { Html = "<p>Nagłówek sekcji 2</p>" } }
        };
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata>(),
            Arg.Any<HeaderFooterContent>(), Arg.Any<HeaderFooterContent>(), Arg.Any<PageMargins>(),
            Arg.Any<PageSize>(), Arg.Any<IReadOnlyList<SectionHeaderFooter>>())
            .Returns(new byte[] { 1, 2 });

        var command = new SaveDocumentCommand(
            Html: "<p>Test</p>",
            OriginalFileName: "test.docx",
            Metadata: null,
            Header: null,
            Footer: null,
            PageSize: pageSize,
            SectionHeadersFooters: sectionHf
        );

        // Act
        var result = await _handler.Handle(command, CancellationToken.None);

        // Assert — lista sekcji musi dojść do konwertera nietknięta, inaczej autosave
        // spłaszcza nagłówki sekcyjne do jednego kompletu.
        result.IsSuccess.Should().BeTrue();
        _converter.Received(1).Convert(
            Arg.Is("<p>Test</p>"),
            Arg.Is<DocumentMetadata?>(m => m == null),
            Arg.Is<HeaderFooterContent?>(h => h == null),
            Arg.Is<HeaderFooterContent?>(f => f == null),
            Arg.Is<PageMargins?>(m => m == null),
            Arg.Is(pageSize),
            Arg.Is<IReadOnlyList<SectionHeaderFooter>?>(s => ReferenceEquals(s, sectionHf)));
    }

    [Test]
    public async Task Handle_WithoutSectionHeadersFooters_ShouldPassNullToConverter()
    {
        // Arrange
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata>(),
            Arg.Any<HeaderFooterContent>(), Arg.Any<HeaderFooterContent>(), Arg.Any<PageMargins>(),
            Arg.Any<PageSize>(), Arg.Any<IReadOnlyList<SectionHeaderFooter>>())
            .Returns(new byte[] { 1 });

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
        _converter.Received(1).Convert(
            Arg.Any<string>(), Arg.Any<DocumentMetadata?>(),
            Arg.Any<HeaderFooterContent?>(), Arg.Any<HeaderFooterContent?>(), Arg.Any<PageMargins?>(),
            Arg.Is<PageSize?>(p => p == null),
            Arg.Is<IReadOnlyList<SectionHeaderFooter>?>(s => s == null));
    }

    // ---------- pass-through (ADR-0031 follow-up): zapis z MasterId zachowuje pakiet oryginału ----------

    private static Document DocxWithBaseVersion(out string storagePath)
    {
        var doc = new Document(Guid.NewGuid(), "doc.docx", DocxMime, "User", metadata: null);
        storagePath = "documents/v1";
        doc.AddVersion(Guid.NewGuid(), storagePath, 123, "User");
        return doc;
    }

    private static SaveDocumentCommand CmdWithMaster(Guid masterId) =>
        new(Html: "<p>Test</p>", OriginalFileName: "test.docx", Metadata: null,
            Header: null, Footer: null, MasterId: masterId);

    [Test]
    public async Task Handle_WithMasterId_AndDocxBaseVersion_UsesPassThrough_NotRegeneration()
    {
        // Arrange — dokument DOCX z wersją bazową: zapis MUSI iść przez ConvertPreservingPackage,
        // żeby styles.xml/theme/numbering oryginału (w tym definicje stylów tabel) przeżyły zapis.
        var repo = Substitute.For<IDocumentRepository>();
        var storage = Substitute.For<IDocumentStorageService>();
        var doc = DocxWithBaseVersion(out var storagePath);
        repo.GetByIdWithVersionsAsync(doc.Id, Arg.Any<CancellationToken>()).Returns(doc);
        var original = new byte[] { 10, 20, 30 };
        storage.DownloadAsync(storagePath, Arg.Any<CancellationToken>()).Returns(original);

        var expected = new byte[] { 7, 7, 7 };
        _converter.ConvertPreservingPackage(
                Arg.Any<string>(), Arg.Any<Stream?>(), Arg.Any<DocumentMetadata?>(),
                Arg.Any<HeaderFooterContent?>(), Arg.Any<HeaderFooterContent?>(),
                Arg.Any<PageMargins?>(), Arg.Any<PageSize?>(), Arg.Any<IReadOnlyList<SectionHeaderFooter>?>())
            .Returns(expected);

        var handler = new SaveDocumentCommandHandler(_converter, repo, storage);

        // Act
        var result = await handler.Handle(CmdWithMaster(doc.Id), CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value!.DocxBytes.Should().Equal(expected);
        await storage.Received(1).DownloadAsync(storagePath, Arg.Any<CancellationToken>());
        _converter.DidNotReceive().Convert(
            Arg.Any<string>(), Arg.Any<DocumentMetadata?>(), Arg.Any<HeaderFooterContent?>(),
            Arg.Any<HeaderFooterContent?>(), Arg.Any<PageMargins?>(), Arg.Any<PageSize?>(),
            Arg.Any<IReadOnlyList<SectionHeaderFooter>?>());
    }

    [Test]
    public async Task Handle_WithMasterId_ButNoBaseVersion_FallsBackToRegeneration()
    {
        var repo = Substitute.For<IDocumentRepository>();
        var storage = Substitute.For<IDocumentStorageService>();
        var doc = new Document(Guid.NewGuid(), "doc.docx", DocxMime, "User", metadata: null); // brak wersji
        repo.GetByIdWithVersionsAsync(doc.Id, Arg.Any<CancellationToken>()).Returns(doc);
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata?>(), Arg.Any<HeaderFooterContent?>(),
            Arg.Any<HeaderFooterContent?>(), Arg.Any<PageMargins?>(), Arg.Any<PageSize?>(),
            Arg.Any<IReadOnlyList<SectionHeaderFooter>?>()).Returns(new byte[] { 1 });

        var handler = new SaveDocumentCommandHandler(_converter, repo, storage);

        var result = await handler.Handle(CmdWithMaster(doc.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        _converter.Received(1).Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata?>(),
            Arg.Any<HeaderFooterContent?>(), Arg.Any<HeaderFooterContent?>(), Arg.Any<PageMargins?>(),
            Arg.Any<PageSize?>(), Arg.Any<IReadOnlyList<SectionHeaderFooter>?>());
        await storage.DidNotReceive().DownloadAsync(Arg.Any<string>(), Arg.Any<CancellationToken>());
    }

    [Test]
    public async Task Handle_WithMasterId_ButNonDocx_FallsBackToRegeneration()
    {
        var repo = Substitute.For<IDocumentRepository>();
        var storage = Substitute.For<IDocumentStorageService>();
        var doc = new Document(Guid.NewGuid(), "scan.pdf", "application/pdf", "User", metadata: null);
        doc.AddVersion(Guid.NewGuid(), "documents/v1", 1, "User");
        repo.GetByIdWithVersionsAsync(doc.Id, Arg.Any<CancellationToken>()).Returns(doc);
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata?>(), Arg.Any<HeaderFooterContent?>(),
            Arg.Any<HeaderFooterContent?>(), Arg.Any<PageMargins?>(), Arg.Any<PageSize?>(),
            Arg.Any<IReadOnlyList<SectionHeaderFooter>?>()).Returns(new byte[] { 1 });

        var handler = new SaveDocumentCommandHandler(_converter, repo, storage);

        var result = await handler.Handle(CmdWithMaster(doc.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        _converter.Received(1).Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata?>(),
            Arg.Any<HeaderFooterContent?>(), Arg.Any<HeaderFooterContent?>(), Arg.Any<PageMargins?>(),
            Arg.Any<PageSize?>(), Arg.Any<IReadOnlyList<SectionHeaderFooter>?>());
        await storage.DidNotReceive().DownloadAsync(Arg.Any<string>(), Arg.Any<CancellationToken>());
    }

    [Test]
    public async Task Handle_PassThroughThrows_FallsBackToRegeneration_BestEffort()
    {
        // Uszkodzony/nieoczekiwany oryginał NIE może wywalić zapisu — best-effort fallback.
        var repo = Substitute.For<IDocumentRepository>();
        var storage = Substitute.For<IDocumentStorageService>();
        var doc = DocxWithBaseVersion(out var storagePath);
        repo.GetByIdWithVersionsAsync(doc.Id, Arg.Any<CancellationToken>()).Returns(doc);
        storage.DownloadAsync(storagePath, Arg.Any<CancellationToken>())
            .Returns<byte[]>(_ => throw new InvalidOperationException("GCS down"));
        var expected = new byte[] { 9 };
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata?>(), Arg.Any<HeaderFooterContent?>(),
            Arg.Any<HeaderFooterContent?>(), Arg.Any<PageMargins?>(), Arg.Any<PageSize?>(),
            Arg.Any<IReadOnlyList<SectionHeaderFooter>?>()).Returns(expected);

        var handler = new SaveDocumentCommandHandler(_converter, repo, storage);

        var result = await handler.Handle(CmdWithMaster(doc.Id), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.DocxBytes.Should().Equal(expected);
    }

    [Test]
    public async Task Handle_MasterIdButNoStorageDeps_FallsBackToRegeneration()
    {
        // Handler skonstruowany bez repo/storage (np. ścieżka „nowy dokument") — MasterId ignorowane.
        _converter.Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata?>(), Arg.Any<HeaderFooterContent?>(),
            Arg.Any<HeaderFooterContent?>(), Arg.Any<PageMargins?>(), Arg.Any<PageSize?>(),
            Arg.Any<IReadOnlyList<SectionHeaderFooter>?>()).Returns(new byte[] { 1 });

        var result = await _handler.Handle(CmdWithMaster(Guid.NewGuid()), CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        _converter.Received(1).Convert(Arg.Any<string>(), Arg.Any<DocumentMetadata?>(),
            Arg.Any<HeaderFooterContent?>(), Arg.Any<HeaderFooterContent?>(), Arg.Any<PageMargins?>(),
            Arg.Any<PageSize?>(), Arg.Any<IReadOnlyList<SectionHeaderFooter>?>());
    }
}
