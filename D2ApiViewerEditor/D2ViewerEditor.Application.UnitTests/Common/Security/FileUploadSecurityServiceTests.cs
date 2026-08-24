using System.IO.Compression;
using D2ViewerEditor.Application.Common.Security;
using FluentAssertions;
using Microsoft.Extensions.Options;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Common.Security;

[TestFixture]
public class FileUploadSecurityServiceTests
{
    [Test]
    public async Task ValidateDocumentAsync_PdfWithBadSignature_ReturnsFailure()
    {
        var sut = BuildService();

        var result = await sut.ValidateDocumentAsync([1, 2, 3], "file.pdf", "application/pdf", CancellationToken.None);

        result.IsValid.Should().BeFalse();
        result.Code.Should().Be(UploadRejectionCode.SignatureMismatch);
    }

    [Test]
    public async Task ValidateDocumentAsync_BinaryDocWithCfbSignature_IsAccepted()
    {
        // Bez wpisu .doc w allowliście KAŻDY upload .doc był odrzucany („Wspierane są tylko
        // pliki DOCX i PDF") — pliki DOC nie istniały w magazynie i na liście admina.
        var sut = BuildService();
        var cfb = new byte[512];
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(cfb, 0);

        var result = await sut.ValidateDocumentAsync(cfb, "stary.doc", "application/msword", CancellationToken.None);

        result.IsValid.Should().BeTrue();
        result.NormalizedMimeType.Should().Be("application/msword");
    }

    [Test]
    public async Task ValidateDocumentAsync_DocWithGarbageBytes_ReturnsSignatureMismatch()
    {
        var sut = BuildService();

        var result = await sut.ValidateDocumentAsync([1, 2, 3, 4], "stary.doc", "application/msword", CancellationToken.None);

        result.IsValid.Should().BeFalse();
        result.Code.Should().Be(UploadRejectionCode.SignatureMismatch);
    }

    [Test]
    public async Task ValidateDocumentAsync_DocThatIsReallyZip_RunsDocxArchiveValidation()
    {
        // „.doc" będący w istocie DOCX (ZIP) — archiwum przechodzi tę samą walidację co .docx.
        var sut = BuildService();
        var zip = BuildDocxArchive("../evil.txt", [1, 2, 3]);

        var result = await sut.ValidateDocumentAsync(zip, "stary.doc", "application/msword", CancellationToken.None);

        result.IsValid.Should().BeFalse();
        result.Code.Should().Be(UploadRejectionCode.DocxZipSlipRisk);
    }

    [Test]
    public async Task ValidateDocumentAsync_DocxWithZipSlipEntry_ReturnsFailure()
    {
        var sut = BuildService();
        var docx = BuildDocxArchive("../evil.txt", [1, 2, 3]);

        var result = await sut.ValidateDocumentAsync(
            docx,
            "file.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            CancellationToken.None);

        result.IsValid.Should().BeFalse();
        result.Code.Should().Be(UploadRejectionCode.DocxZipSlipRisk);
    }

    [Test]
    public async Task ValidateDocumentAsync_DocxWithMacroProject_ReturnsFailure()
    {
        var sut = BuildService();
        var docx = BuildDocxArchive("word/vbaProject.bin", [1, 2, 3]);

        var result = await sut.ValidateDocumentAsync(
            docx,
            "file.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            CancellationToken.None);

        result.IsValid.Should().BeFalse();
        result.Code.Should().Be(UploadRejectionCode.DocxSuspiciousEntry);
    }

    [Test]
    public async Task ValidateImageAsync_ExtensionMimeMismatch_ReturnsFailure()
    {
        var sut = BuildService();
        var pngHeader = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

        var result = await sut.ValidateImageAsync(
            pngHeader,
            "image.jpg",
            "image/png",
            CancellationToken.None);

        result.IsValid.Should().BeFalse();
        result.Code.Should().Be(UploadRejectionCode.ExtensionMimeMismatch);
    }

    [Test]
    public async Task ValidateImageAsync_WhenScannerReportsInfected_ReturnsFailure()
    {
        var scanner = new StubScanner(_ => new FileScanResult(FileScanStatus.Infected, "EICAR"));
        var sut = BuildService(scanner: scanner);
        var jpgHeader = new byte[] { 0xFF, 0xD8, 0xFF, 0xE0 };

        var result = await sut.ValidateImageAsync(jpgHeader, "photo.jpg", "image/jpeg", CancellationToken.None);

        result.IsValid.Should().BeFalse();
        result.Code.Should().Be(UploadRejectionCode.MalwareDetected);
    }

    [Test]
    public async Task ValidateImageAsync_WhenScannerUnavailableAndFailClosed_ReturnsFailure()
    {
        var scanner = new StubScanner(_ => new FileScanResult(FileScanStatus.Unavailable));
        var sut = BuildService(scanner: scanner);
        var jpgHeader = new byte[] { 0xFF, 0xD8, 0xFF, 0xE0 };

        var result = await sut.ValidateImageAsync(jpgHeader, "photo.jpg", "image/jpeg", CancellationToken.None);

        result.IsValid.Should().BeFalse();
        result.Code.Should().Be(UploadRejectionCode.MalwareScanUnavailable);
    }

    private static FileUploadSecurityService BuildService(
        UploadSecurityOptions? options = null,
        IFileScanner? scanner = null)
    {
        return new FileUploadSecurityService(
            Options.Create(options ?? new UploadSecurityOptions()),
            scanner ?? new NoOpFileScanner());
    }

    private static byte[] BuildDocxArchive(string? extraEntryName = null, byte[]? extraEntryContent = null)
    {
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "[Content_Types].xml", "<Types/>");
            WriteEntry(archive, "_rels/.rels", "<Relationships/>");
            WriteEntry(archive, "word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>");

            if (!string.IsNullOrWhiteSpace(extraEntryName))
            {
                var payload = extraEntryContent ?? [0x01];
                WriteEntry(archive, extraEntryName, payload);
            }
        }

        return ms.ToArray();
    }

    private static void WriteEntry(ZipArchive archive, string name, string value)
    {
        var entry = archive.CreateEntry(name);
        using var stream = entry.Open();
        stream.FillText(value);
    }

    private static void WriteEntry(ZipArchive archive, string name, byte[] content)
    {
        var entry = archive.CreateEntry(name);
        using var stream = entry.Open();
        stream.Fill(content);
    }

    private sealed class StubScanner : IFileScanner
    {
        private readonly Func<FileScanRequest, FileScanResult> _scan;

        public StubScanner(Func<FileScanRequest, FileScanResult> scan)
        {
            _scan = scan;
        }

        public Task<FileScanResult> ScanAsync(FileScanRequest request, CancellationToken cancellationToken)
        {
            return Task.FromResult(_scan(request));
        }
    }
}
