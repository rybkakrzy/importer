using D2ViewerEditor.Application.Features.Documents.Commands.IngestExternalDocument;
using D2ViewerEditor.Application.Common.Security;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Documents.Commands;

[TestFixture]
public class IngestExternalDocumentCommandHandlerTests
{
    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    private const string PdfMime = "application/pdf";

    private IDocumentRepository _repo = null!;
    private IDocumentStorageService _storage = null!;
    private IFileUploadSecurityService _uploadSecurityService = null!;
    private IReturnUrlValidator _returnUrlValidator = null!;
    private IngestExternalDocumentCommandHandler _handler = null!;

    [SetUp]
    public void SetUp()
    {
        _repo = Substitute.For<IDocumentRepository>();
        _storage = Substitute.For<IDocumentStorageService>();
        _uploadSecurityService = Substitute.For<IFileUploadSecurityService>();
        _returnUrlValidator = Substitute.For<IReturnUrlValidator>();
        _storage.UploadAsync(Arg.Any<Guid>(), Arg.Any<byte[]>(), Arg.Any<string>(), Arg.Any<CancellationToken>())
            .Returns(ci => $"documents/{ci.ArgAt<Guid>(0)}");
        _uploadSecurityService.ValidateDocumentAsync(
                Arg.Any<byte[]>(),
                Arg.Any<string>(),
                Arg.Any<string>(),
                Arg.Any<CancellationToken>())
            .Returns(UploadValidationResult.Success(DocxMime));
        _returnUrlValidator.Validate(Arg.Any<string>())
            .Returns(ci => ReturnUrlValidationResult.Success(ci.Arg<string>()));
        _handler = new IngestExternalDocumentCommandHandler(_repo, _storage, _uploadSecurityService, _returnUrlValidator);
    }

    [Test]
    public async Task Handle_Docx_CreatesOriginalPlusEditableVersion_AndReturnsVersionId()
    {
        var cmd = new IngestExternalDocumentCommand(
            new byte[] { 1, 2, 3 },
            "ext.docx",
            DocxMime,
            "ExternalApp",
            "{\"returnUrl\":\"https://example.com/cb\"}");

        var result = await _handler.Handle(cmd, CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.VersionId.Should().NotBeNull();           // DOCX → wersja edytowalna
        result.Value.FileName.Should().Be("ext.docx");
        await _storage.Received(2).UploadAsync(Arg.Any<Guid>(), Arg.Any<byte[]>(), DocxMime, Arg.Any<CancellationToken>());
        await _repo.Received(1).AddAsync(
            Arg.Is<Document>(d => d.Versions.Count == 2), Arg.Any<CancellationToken>());
        await _repo.Received(1).SaveChangesAsync(Arg.Any<CancellationToken>());
    }

    [Test]
    public async Task Handle_Pdf_CreatesSingleVersion_AndVersionIdIsNull()
    {
        var cmd = new IngestExternalDocumentCommand(new byte[] { 1, 2, 3 }, "ext.pdf", PdfMime, "ExternalApp");

        var result = await _handler.Handle(cmd, CancellationToken.None);

        result.IsSuccess.Should().BeTrue();
        result.Value!.VersionId.Should().BeNull();              // PDF → tylko oryginał
        await _storage.Received(1).UploadAsync(Arg.Any<Guid>(), Arg.Any<byte[]>(), PdfMime, Arg.Any<CancellationToken>());
        await _repo.Received(1).AddAsync(
            Arg.Is<Document>(d => d.Versions.Count == 1), Arg.Any<CancellationToken>());
    }

    [Test]
    public async Task Handle_EmptyContent_ReturnsFailure()
    {
        var result = await _handler.Handle(
            new IngestExternalDocumentCommand(Array.Empty<byte>(), "x.docx", DocxMime, "App"),
            CancellationToken.None);

        result.IsFailure.Should().BeTrue();
        await _repo.DidNotReceive().AddAsync(Arg.Any<Document>(), Arg.Any<CancellationToken>());
    }

    [Test]
    public async Task Handle_UnsupportedMime_ReturnsFailure()
    {
        var result = await _handler.Handle(
            new IngestExternalDocumentCommand(new byte[] { 1 }, "x.txt", "text/plain", "App"),
            CancellationToken.None);

        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("Nieobsługiwany");
        await _repo.DidNotReceive().AddAsync(Arg.Any<Document>(), Arg.Any<CancellationToken>());
    }

    [Test]
    public async Task Handle_RepositoryThrows_ReturnsFailure()
    {
        _repo.When(r => r.SaveChangesAsync(Arg.Any<CancellationToken>()))
            .Do(_ => throw new Exception("db error"));

        var result = await _handler.Handle(
            new IngestExternalDocumentCommand(new byte[] { 1 }, "x.pdf", PdfMime, "App"),
            CancellationToken.None);

        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("db error");
    }

    [Test]
    public async Task Handle_DocxWithoutReturnUrl_ReturnsFailure()
    {
        var result = await _handler.Handle(
            new IngestExternalDocumentCommand(new byte[] { 1, 2, 3 }, "ext.docx", DocxMime, "ExternalApp"),
            CancellationToken.None);

        result.IsFailure.Should().BeTrue();
        result.Error.Should().Contain("returnUrl");
    }
}
