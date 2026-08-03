using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Services;

namespace D2ViewerEditor.Benchmarks.Benchmarks;

/// <summary>
/// Benchmarki podpisywania i weryfikacji dokumentów DOCX (DigitalSignatureService).
///
/// Mierzone obszary:
///  - SignDocument: SHA-256 haszowania MainDocumentPart + RSA-PKCS1 podpisanie + budowanie XML + zapis Custom XML Part
///  - VerifySignatures: ponowne haszowanie + RSA weryfikacja dla każdego podpisu
///  - Wpływ liczby podpisów na koszt weryfikacji (O(N) per podpis)
///
/// Certyfikat jest generowany programowo (self-signed RSA-2048) w GlobalSetup.
/// Benchmarki używają prawdziwego kryptograficzno-XML pipeline — żadnych mocków.
/// </summary>
[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
public class DigitalSignatureBenchmarks
{
    // Ephemeral, per-run passphrase for the in-memory self-signed PFX — never a stored/hardcoded credential.
    private static readonly string CertPassword = Convert.ToBase64String(RandomNumberGenerator.GetBytes(32));

    private DigitalSignatureService _service = null!;
    private byte[] _certBytes = null!;
    private byte[] _baseDocx = null!;
    private byte[] _signedOnce = null!;
    private byte[] _signedThreeTimes = null!;

    [GlobalSetup]
    public void Setup()
    {
        _service = new DigitalSignatureService();

        // Wygeneruj self-signed RSA-2048 cert (bez wymogu zewnętrznego pliku PFX)
        using var rsa = RSA.Create(2048);
        var request = new CertificateRequest(
            "CN=D2ViewerEditor Benchmark Signer, O=Benchmark Suite, C=PL",
            rsa,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);

        request.CertificateExtensions.Add(
            new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));

        using var cert = request.CreateSelfSigned(
            DateTimeOffset.UtcNow.AddDays(-1),
            DateTimeOffset.UtcNow.AddYears(2));

        _certBytes = cert.Export(X509ContentType.Pfx, CertPassword);

        // Przygotuj bazowy dokument DOCX przez produkcyjny konwerter
        var htmlConverter = new HtmlToDocxConverter();
        _baseDocx = htmlConverter.Convert(
            "<h1>Dokument do podpisania</h1><p>Treść dokumentu podlegającego podpisowi cyfrowemu.</p><p>Dane kontrahenta: Firma XYZ Sp. z o.o., NIP: 123-456-78-90.</p>");

        // Przygotuj dokumenty pre-signed do benchmarku weryfikacji
        _signedOnce = Sign(_baseDocx, "Podpisujący Pierwszy", "Dyrektor");

        _signedThreeTimes = Sign(
            Sign(
                Sign(_baseDocx, "Podpisujący Pierwszy", "Prezes"),
                "Podpisujący Drugi", "Dyrektor Finansowy"),
            "Podpisujący Trzeci", "Kierownik Projektu");
    }

    private byte[] Sign(byte[] docx, string signerName, string? signerTitle)
        => _service.SignDocument(docx, _certBytes, CertPassword, signerName, signerTitle,
            $"{signerName.Replace(" ", ".").ToLowerInvariant()}@example.com");

    /// <summary>
    /// Podpisuje dokument DOCX jednym podpisem.
    /// Pipeline: załaduj PFX → SHA-256 hash MainDocumentPart → RSA sign → zapis Custom XML Part.
    /// </summary>
    [Benchmark(Baseline = true)]
    public byte[] SignDocument()
        => _service.SignDocument(
            _baseDocx, _certBytes, CertPassword,
            "Benchmark Signer", "Tester", "benchmark@example.com");

    /// <summary>
    /// Weryfikuje jeden podpis w dokumencie.
    /// Pipeline: re-hash MainDocumentPart → RSA verify → porównanie hashy.
    /// Koszt: O(1) podpisów × (SHA-256 całego dokumentu + RSA verify).
    /// </summary>
    [Benchmark]
    public List<DigitalSignatureInfo> VerifySingleSignature()
    {
        using var stream = new MemoryStream(_signedOnce);
        return _service.VerifySignatures(stream);
    }

    /// <summary>
    /// Weryfikuje trzy podpisy w dokumencie.
    /// Haszowanie jest wykonywane raz, weryfikacja RSA — N razy (N = liczba podpisów).
    /// Sprawdza, czy implementacja nie haszuje dokumentu wielokrotnie (powinno być O(1) hash + O(N) verify).
    /// </summary>
    [Benchmark]
    public List<DigitalSignatureInfo> VerifyThreeSignatures()
    {
        using var stream = new MemoryStream(_signedThreeTimes);
        return _service.VerifySignatures(stream);
    }

    /// <summary>
    /// Dodaje drugi podpis do już podpisanego dokumentu.
    /// Ciekawy przypadek: Custom XML Part z pierwszym podpisem musi zostać odczytana,
    /// nowy element dodany i cała część zastąpiona.
    /// </summary>
    [Benchmark]
    public byte[] AddSecondSignatureToSignedDocument()
        => _service.SignDocument(
            _signedOnce, _certBytes, CertPassword,
            "Drugi Podpisujący", "Wiceprezes", "wiceprezes@example.com");
}
