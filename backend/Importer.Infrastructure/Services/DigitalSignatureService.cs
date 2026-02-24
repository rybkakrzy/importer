using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Importer.Domain.Interfaces;
using Importer.Domain.Models;

namespace Importer.Infrastructure.Services;

/// <summary>
/// Serwis do obsługi podpisów cyfrowych w dokumentach DOCX
/// Podpisy są przechowywane jako Custom XML Part w pakiecie OPC
/// </summary>
public class DigitalSignatureService : IDigitalSignatureService
{
    private static readonly XNamespace SigNs = "http://schemas.importer.app/digitalsignatures";
    private const string SignatureContentType = "application/xml";

    /// <summary>
    /// Podpisuje dokument DOCX certyfikatem X.509 (PFX/P12)
    /// </summary>
    public byte[] SignDocument(byte[] docxBytes, byte[] certificateBytes, string password,
                                string signerName, string? signerTitle = null,
                                string? signerEmail = null, string? reason = null)
    {
        // Wczytaj certyfikat
        var cert = new X509Certificate2(certificateBytes, password,
            X509KeyStorageFlags.UserKeySet | X509KeyStorageFlags.Exportable);

        using var memoryStream = new MemoryStream();
        memoryStream.Write(docxBytes, 0, docxBytes.Length);
        memoryStream.Position = 0;

        using (var document = WordprocessingDocument.Open(memoryStream, true))
        {
            // Oblicz hash głównej części dokumentu
            var contentHash = ComputeDocumentHash(document);

            // Podpisz hash kluczem prywatnym
            using var rsa = cert.GetRSAPrivateKey()
                ?? throw new InvalidOperationException("Certyfikat nie zawiera klucza prywatnego RSA. Upewnij się, że plik PFX zawiera klucz prywatny.");

            var signatureBytes = rsa.SignData(contentHash, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

            // Zbuduj XML podpisu
            var signatureXml = BuildSignatureXml(
                signerName, signerTitle, signerEmail, reason,
                cert, contentHash, signatureBytes);

            // Wczytaj istniejące podpisy
            var existingSignatures = ReadSignaturesXml(document);
            existingSignatures.Add(signatureXml);

            // Zapisz podpisy jako Custom XML Part
            SaveSignaturesToDocument(document, existingSignatures);

            document.Save();
        }

        return memoryStream.ToArray();
    }

    /// <summary>
    /// Weryfikuje podpisy cyfrowe w dokumencie DOCX
    /// </summary>
    public List<DigitalSignatureInfo> VerifySignatures(Stream docxStream)
    {
        var signatures = new List<DigitalSignatureInfo>();

        using var memoryStream = new MemoryStream();
        docxStream.CopyTo(memoryStream);
        memoryStream.Position = 0;

        using var document = WordprocessingDocument.Open(memoryStream, false);

        var signatureElements = ReadSignaturesXml(document);
        var currentHash = ComputeDocumentHash(document);

        foreach (var sigEl in signatureElements)
        {
            var info = ParseSignatureElement(sigEl);

            // Weryfikuj podpis
            var storedHash = Convert.FromBase64String(sigEl.Element(SigNs + "DocumentHash")?.Value ?? "");
            var signatureValue = Convert.FromBase64String(sigEl.Element(SigNs + "SignatureValue")?.Value ?? "");
            var certBase64 = sigEl.Element(SigNs + "CertificateData")?.Value;

            if (certBase64 != null && storedHash.Length > 0 && signatureValue.Length > 0)
            {
                try
                {
                    var cert = new X509Certificate2(Convert.FromBase64String(certBase64));
                    using var rsa = cert.GetRSAPublicKey();

                    if (rsa != null)
                    {
                        // Sprawdź czy podpis jest kryptograficznie poprawny
                        var signatureValid = rsa.VerifyData(storedHash, signatureValue,
                            HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

                        // Sprawdź czy dokument nie został zmodyfikowany po podpisaniu
                        var hashMatch = storedHash.SequenceEqual(currentHash);

                        if (signatureValid && hashMatch)
                        {
                            info.IsValid = true;
                            info.ValidationMessage = "Podpis jest prawidłowy. Dokument nie został zmodyfikowany.";
                        }
                        else if (signatureValid && !hashMatch)
                        {
                            info.IsValid = false;
                            info.ValidationMessage = "Podpis kryptograficzny prawidłowy, ale dokument został zmodyfikowany po podpisaniu.";
                        }
                        else
                        {
                            info.IsValid = false;
                            info.ValidationMessage = "Podpis kryptograficzny jest nieprawidłowy.";
                        }
                    }
                    else
                    {
                        info.IsValid = false;
                        info.ValidationMessage = "Nie można zweryfikować podpisu — brak klucza publicznego RSA.";
                    }
                }
                catch (Exception ex)
                {
                    info.IsValid = false;
                    info.ValidationMessage = $"Błąd weryfikacji: {ex.Message}";
                }
            }
            else
            {
                info.IsValid = false;
                info.ValidationMessage = "Niekompletne dane podpisu.";
            }

            signatures.Add(info);
        }

        return signatures;
    }

    /// <summary>
    /// Oblicza SHA-256 hash głównej części dokumentu (treść)
    /// </summary>
    private byte[] ComputeDocumentHash(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart;
        if (mainPart == null)
            return SHA256.HashData(Array.Empty<byte>());

        using var stream = mainPart.GetStream(FileMode.Open, FileAccess.Read);
        return SHA256.HashData(stream);
    }

    /// <summary>
    /// Buduje element XML podpisu
    /// </summary>
    private XElement BuildSignatureXml(
        string signerName, string? signerTitle, string? signerEmail,
        string? reason, X509Certificate2 cert,
        byte[] documentHash, byte[] signatureBytes)
    {
        return new XElement(SigNs + "Signature",
            new XElement(SigNs + "SignerName", signerName),
            new XElement(SigNs + "SignerTitle", signerTitle ?? ""),
            new XElement(SigNs + "SignerEmail", signerEmail ?? ""),
            new XElement(SigNs + "Reason", reason ?? ""),
            new XElement(SigNs + "SignedAt", DateTime.UtcNow.ToString("O")),
            new XElement(SigNs + "CertificateSubject", cert.Subject),
            new XElement(SigNs + "CertificateIssuer", cert.Issuer),
            new XElement(SigNs + "CertificateSerial", cert.SerialNumber),
            new XElement(SigNs + "CertificateValidFrom", cert.NotBefore.ToString("O")),
            new XElement(SigNs + "CertificateValidTo", cert.NotAfter.ToString("O")),
            new XElement(SigNs + "DocumentHash", Convert.ToBase64String(documentHash)),
            new XElement(SigNs + "SignatureValue", Convert.ToBase64String(signatureBytes)),
            new XElement(SigNs + "CertificateData", Convert.ToBase64String(cert.RawData))
        );
    }

    /// <summary>
    /// Parsuje element XML podpisu na DigitalSignatureInfo
    /// </summary>
    private DigitalSignatureInfo ParseSignatureElement(XElement element)
    {
        return new DigitalSignatureInfo
        {
            SignerName = element.Element(SigNs + "SignerName")?.Value ?? "",
            SignerTitle = element.Element(SigNs + "SignerTitle")?.Value,
            SignerEmail = element.Element(SigNs + "SignerEmail")?.Value,
            Reason = element.Element(SigNs + "Reason")?.Value,
            CertificateSubject = element.Element(SigNs + "CertificateSubject")?.Value ?? "",
            CertificateIssuer = element.Element(SigNs + "CertificateIssuer")?.Value ?? "",
            CertificateSerialNumber = element.Element(SigNs + "CertificateSerial")?.Value ?? "",
            SignedAt = DateTime.TryParse(element.Element(SigNs + "SignedAt")?.Value, out var signedAt) ? signedAt : DateTime.MinValue,
            CertificateValidFrom = DateTime.TryParse(element.Element(SigNs + "CertificateValidFrom")?.Value, out var from) ? from : DateTime.MinValue,
            CertificateValidTo = DateTime.TryParse(element.Element(SigNs + "CertificateValidTo")?.Value, out var to) ? to : DateTime.MaxValue,
        };
    }

    /// <summary>
    /// Odczytuje istniejące elementy podpisów z Custom XML Part
    /// </summary>
    private List<XElement> ReadSignaturesXml(WordprocessingDocument document)
    {
        var result = new List<XElement>();

        if (document.MainDocumentPart == null) return result;

        foreach (var xmlPart in document.MainDocumentPart.CustomXmlParts)
        {
            try
            {
                using var stream = xmlPart.GetStream(FileMode.Open, FileAccess.Read);
                var doc = XDocument.Load(stream);

                if (doc.Root?.Name.Namespace == SigNs && doc.Root.Name.LocalName == "DigitalSignatures")
                {
                    result.AddRange(doc.Root.Elements(SigNs + "Signature"));
                }
            }
            catch
            {
                // Ignoruj nieprawidłowe XML parts
            }
        }

        return result;
    }

    /// <summary>
    /// Zapisuje podpisy jako Custom XML Part w dokumencie
    /// </summary>
    private void SaveSignaturesToDocument(WordprocessingDocument document, List<XElement> signatures)
    {
        if (document.MainDocumentPart == null) return;

        // Usuń istniejący part z podpisami
        var partsToRemove = new List<CustomXmlPart>();
        foreach (var xmlPart in document.MainDocumentPart.CustomXmlParts)
        {
            try
            {
                using var stream = xmlPart.GetStream(FileMode.Open, FileAccess.Read);
                var doc = XDocument.Load(stream);
                if (doc.Root?.Name.Namespace == SigNs)
                {
                    partsToRemove.Add(xmlPart);
                }
            }
            catch { }
        }

        foreach (var part in partsToRemove)
        {
            document.MainDocumentPart.DeletePart(part);
        }

        // Utwórz nowy part z podpisami
        if (signatures.Count > 0)
        {
            var xmlDoc = new XDocument(
                new XElement(SigNs + "DigitalSignatures", signatures));

            var newPart = document.MainDocumentPart.AddCustomXmlPart(SignatureContentType);
            using var writeStream = newPart.GetStream(FileMode.Create, FileAccess.Write);
            xmlDoc.Save(writeStream);
        }
    }
}
