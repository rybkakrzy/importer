using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using D2ViewerEditor.Infrastructure.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Podpis cyfrowy DOCX (Custom XML Part, X.509 + RSA-SHA256): round-trip sign→verify,
/// wykrycie modyfikacji dokumentu po podpisaniu, dokument bez podpisów.
/// </summary>
[TestFixture]
public class DigitalSignatureServiceTests
{
    // Ephemeral, per-run passphrase for the throwaway PFX — never a stored/hardcoded credential.
    private static readonly string CertPassword = Convert.ToBase64String(RandomNumberGenerator.GetBytes(32));

    private DigitalSignatureService _svc = null!;

    [SetUp]
    public void SetUp() => _svc = new DigitalSignatureService();

    [Test]
    public void SignThenVerify_ValidSignature_IsValidTrue()
    {
        var pfx = CreateSelfSignedPfx(CertPassword);
        var docx = MinimalDocx("Treść do podpisania");

        var signed = _svc.SignDocument(docx, pfx, CertPassword, "Jan Kowalski", "Dyrektor", "jan@example.com", "Zatwierdzenie");

        var infos = _svc.VerifySignatures(new MemoryStream(signed));
        infos.Should().HaveCount(1);
        infos[0].SignerName.Should().Be("Jan Kowalski");
        infos[0].IsValid.Should().BeTrue();
    }

    [Test]
    public void Verify_AfterTampering_IsValidFalse()
    {
        var pfx = CreateSelfSignedPfx(CertPassword);
        var signed = _svc.SignDocument(MinimalDocx("Oryginał"), pfx, CertPassword, "Jan Kowalski");

        var tampered = AppendParagraph(signed, "Doklejone po podpisie");

        var infos = _svc.VerifySignatures(new MemoryStream(tampered));
        infos.Should().HaveCount(1);
        infos[0].IsValid.Should().BeFalse(); // hash treści ≠ hash zapisany w podpisie
    }

    [Test]
    public void Verify_DocumentWithoutSignatures_ReturnsEmpty()
    {
        var infos = _svc.VerifySignatures(new MemoryStream(MinimalDocx("Bez podpisu")));

        infos.Should().BeEmpty();
    }

    // ---- helpers ----------------------------------------------------------------

    private static byte[] CreateSelfSignedPfx(string password)
    {
        using var rsa = RSA.Create(2048);
        var req = new CertificateRequest("CN=Test Signer", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        using var cert = req.CreateSelfSigned(DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddYears(1));
        return cert.Export(X509ContentType.Pfx, password);
    }

    private static byte[] MinimalDocx(string text)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static byte[] AppendParagraph(byte[] docx, string text)
    {
        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            doc.MainDocumentPart!.Document.Body!.AppendChild(new Paragraph(new Run(new Text(text))));
            doc.MainDocumentPart.Document.Save();
        }
        return ms.ToArray();
    }
}
