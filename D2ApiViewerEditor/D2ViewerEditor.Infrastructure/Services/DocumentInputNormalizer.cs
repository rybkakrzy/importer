using D2ViewerEditor.Domain.Interfaces;
using NPOI.POIFS.FileSystem;

namespace D2ViewerEditor.Infrastructure.Services;

/// <summary>
/// Doprowadza wejściowy plik do zwykłego DOCX (ZIP/OOXML) przed parsowaniem:
/// - ZIP (PK) → już DOCX (obejmuje też .doc będący w istocie mislabeled DOCX),
/// - CFB (compound file) z EncryptionInfo → DOCX zaszyfrowany hasłem → dekrypcja (NPOI),
/// - CFB z WordDocument → binarny .doc → brak wbudowanego konwertera (kontrolowany status).
/// Czysto zarządzane (NPOI POIFS/Crypt) — bez LibreOffice/System.Drawing (Linux/GCP-safe).
/// </summary>
public sealed class DocumentInputNormalizer : IDocumentInputNormalizer
{
    // Magic bytes pakietu ZIP (DOCX) i kontenera CFB (OLE/CFBF: .doc lub zaszyfrowany OOXML).
    private static readonly byte[] ZipMagic = { 0x50, 0x4B, 0x03, 0x04 };
    private static readonly byte[] CfbMagic = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    public DocumentInputResult Normalize(byte[] bytes, string? password = null)
    {
        if (bytes == null || bytes.Length < 8)
            return DocumentInputResult.Failure(DocumentInputStatus.Invalid);

        // 1. Zwykły DOCX (ZIP). Także .doc, które naprawdę jest DOCX (częsty mislabeled) → pass-through.
        if (StartsWith(bytes, ZipMagic))
            return DocumentInputResult.Success(bytes);

        // 2. Kontener CFB — albo zaszyfrowany OOXML, albo binarny .doc.
        if (!StartsWith(bytes, CfbMagic))
            return DocumentInputResult.Failure(DocumentInputStatus.Invalid);

        POIFSFileSystem fs;
        try
        {
            fs = new POIFSFileSystem(new MemoryStream(bytes, writable: false));
        }
        catch
        {
            return DocumentInputResult.Failure(DocumentInputStatus.Invalid);
        }

        // 2a. Zaszyfrowany OOXML (DOCX z hasłem) — strumienie EncryptionInfo + EncryptedPackage.
        // NPOI czyta kontener CFB, ale NIE dekryptuje Agile (brak implementacji) → robimy to sami
        // (OoxmlAgileDecryptor). NPOI tu tylko wydobywa surowe strumienie.
        if (fs.Root.HasEntry("EncryptionInfo") && fs.Root.HasEntry("EncryptedPackage"))
        {
            if (string.IsNullOrEmpty(password))
                return DocumentInputResult.Failure(DocumentInputStatus.PasswordRequired);

            byte[] encryptionInfo, encryptedPackage;
            try
            {
                encryptionInfo = ReadCfbStream(fs, "EncryptionInfo");
                encryptedPackage = ReadCfbStream(fs, "EncryptedPackage");
            }
            catch
            {
                return DocumentInputResult.Failure(DocumentInputStatus.Invalid);
            }

            var result = OoxmlAgileDecryptor.TryDecrypt(encryptionInfo, encryptedPackage, password, out var docx);
            return result switch
            {
                OoxmlAgileDecryptor.DecryptResult.Ok => DocumentInputResult.Success(docx!),
                OoxmlAgileDecryptor.DecryptResult.WrongPassword => DocumentInputResult.Failure(DocumentInputStatus.WrongPassword),
                // Standard/CryptoAPI (Office 2007) nieobsługiwane — sygnalizujemy jak nierozpoznane.
                OoxmlAgileDecryptor.DecryptResult.Unsupported => DocumentInputResult.Failure(DocumentInputStatus.Invalid),
                _ => DocumentInputResult.Failure(DocumentInputStatus.Invalid)
            };
        }

        // 2b. Binarny .doc (starszy Word) — strumień WordDocument. Brak wbudowanego, czysto zarządzanego
        // konwertera .doc→.docx (NPOI nie ma HWPF; b2xtranslator NuGet nie zawiera parsera DocFileFormat;
        // LibreOffice/System.Drawing zakazane). Zwracamy kontrolowany status — bez udawania konwersji.
        if (fs.Root.HasEntry("WordDocument"))
            return DocumentInputResult.Failure(DocumentInputStatus.UnsupportedLegacyDoc);

        return DocumentInputResult.Failure(DocumentInputStatus.Invalid);
    }

    private static byte[] ReadCfbStream(POIFSFileSystem fs, string name)
    {
        using var input = fs.CreateDocumentInputStream(name);
        using var ms = new MemoryStream();
        input.CopyTo(ms);
        return ms.ToArray();
    }

    private static bool StartsWith(byte[] data, byte[] prefix)
    {
        if (data.Length < prefix.Length) return false;
        for (int i = 0; i < prefix.Length; i++)
            if (data[i] != prefix[i]) return false;
        return true;
    }
}
