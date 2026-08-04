using System.IO.Compression;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Application.Common.Security;

public sealed class FileUploadSecurityService : IFileUploadSecurityService
{
    private static readonly Dictionary<string, string> DocumentMimeByExtension = new(StringComparer.OrdinalIgnoreCase)
    {
        [".docx"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        // Binarny .doc jest wspierany (LegacyDocBinaryConverter) — bez tego wpisu KAŻDY
        // upload .doc z dashboardu był odrzucany i pliki DOC nie istniały w magazynie
        // (admin „Lista plików" pokazywała wyłącznie DOCX/PDF).
        [".doc"] = "application/msword",
        [".pdf"] = "application/pdf"
    };

    private static readonly Dictionary<string, string> ImageMimeByExtension = new(StringComparer.OrdinalIgnoreCase)
    {
        [".jpg"] = "image/jpeg",
        [".jpeg"] = "image/jpeg",
        [".png"] = "image/png",
        [".gif"] = "image/gif",
        [".bmp"] = "image/bmp",
        [".webp"] = "image/webp"
    };

    private static readonly HashSet<string> SuspiciousDocxExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".exe", ".dll", ".bat", ".cmd", ".ps1", ".vbs", ".js", ".scr", ".com"
    };

    private readonly UploadSecurityOptions _options;
    private readonly IFileScanner _fileScanner;

    public FileUploadSecurityService(IOptions<UploadSecurityOptions> options, IFileScanner fileScanner)
    {
        _options = options.Value;
        _fileScanner = fileScanner;
    }

    public async Task<UploadValidationResult> ValidateDocumentAsync(
        byte[] content,
        string fileName,
        string mimeType,
        CancellationToken cancellationToken)
    {
        if (content.Length == 0)
            return UploadValidationResult.Failure(UploadRejectionCode.Empty, "Zawartość dokumentu nie może być pusta.");

        var extension = Path.GetExtension(fileName);
        if (!DocumentMimeByExtension.TryGetValue(extension, out var expectedMime))
            return UploadValidationResult.Failure(UploadRejectionCode.UnsupportedExtension,
                "Wspierane są tylko pliki DOCX i PDF.");

        if (!string.Equals(expectedMime, mimeType, StringComparison.OrdinalIgnoreCase))
            return UploadValidationResult.Failure(UploadRejectionCode.ExtensionMimeMismatch,
                "Rozszerzenie pliku nie zgadza się z deklarowanym typem MIME.");

        if (string.Equals(expectedMime, "application/pdf", StringComparison.Ordinal))
        {
            if (!LooksLikePdf(content))
                return UploadValidationResult.Failure(UploadRejectionCode.SignatureMismatch,
                    "Podpis binarny pliku nie odpowiada PDF.");
        }
        else if (string.Equals(expectedMime, "application/msword", StringComparison.Ordinal))
        {
            // Binarny .doc = kontener CFB (D0 CF 11 E0…). Częsty wariant to też „.doc" będący
            // w istocie DOCX (ZIP) — wtedy walidujemy archiwum jak dla .docx; normalizer
            // otwarcia i tak rozpozna format po zawartości.
            if (LooksLikeZip(content))
            {
                var docAsDocxCheck = ValidateDocxArchive(content);
                if (!docAsDocxCheck.IsValid)
                    return docAsDocxCheck;
            }
            else if (!LooksLikeCfb(content))
            {
                return UploadValidationResult.Failure(UploadRejectionCode.SignatureMismatch,
                    "Podpis binarny pliku nie odpowiada DOC.");
            }
        }
        else
        {
            if (!LooksLikeZip(content))
                return UploadValidationResult.Failure(UploadRejectionCode.SignatureMismatch,
                    "Podpis binarny pliku nie odpowiada DOCX.");

            var archiveCheck = ValidateDocxArchive(content);
            if (!archiveCheck.IsValid)
                return archiveCheck;
        }

        var scan = await _fileScanner.ScanAsync(new FileScanRequest(fileName, expectedMime, content), cancellationToken);
        var scanFailure = EvaluateScanFailure(scan);
        return scanFailure ?? UploadValidationResult.Success(expectedMime);
    }

    public async Task<UploadValidationResult> ValidateImageAsync(
        byte[] content,
        string fileName,
        string contentType,
        CancellationToken cancellationToken)
    {
        if (content.Length == 0)
            return UploadValidationResult.Failure(UploadRejectionCode.Empty, "Plik obrazu nie może być pusty.");

        var extension = Path.GetExtension(fileName);
        if (!ImageMimeByExtension.TryGetValue(extension, out var expectedMime))
            return UploadValidationResult.Failure(UploadRejectionCode.UnsupportedExtension,
                "Niedozwolony format obrazu.");

        if (!string.Equals(expectedMime, contentType, StringComparison.OrdinalIgnoreCase))
            return UploadValidationResult.Failure(UploadRejectionCode.ExtensionMimeMismatch,
                "Rozszerzenie pliku nie zgadza się z deklarowanym typem obrazu.");

        if (!MatchesImageSignature(expectedMime, content))
            return UploadValidationResult.Failure(UploadRejectionCode.SignatureMismatch,
                "Podpis binarny obrazu nie zgadza się z deklarowanym formatem.");

        var scan = await _fileScanner.ScanAsync(new FileScanRequest(fileName, expectedMime, content), cancellationToken);
        var scanFailure = EvaluateScanFailure(scan);
        return scanFailure ?? UploadValidationResult.Success(expectedMime);
    }

    public UploadValidationResult ValidateDocxStructure(byte[] content)
    {
        if (content.Length == 0)
            return UploadValidationResult.Failure(UploadRejectionCode.Empty, "Zawartość dokumentu nie może być pusta.");

        if (!LooksLikeZip(content))
            return UploadValidationResult.Failure(UploadRejectionCode.SignatureMismatch,
                "Podpis binarny pliku nie odpowiada DOCX.");

        return ValidateDocxArchive(content);
    }

    private UploadValidationResult ValidateDocxArchive(byte[] content)
    {
        try
        {
            using var ms = new MemoryStream(content, writable: false);
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: false);

            if (archive.Entries.Count > _options.MaxDocxEntryCount)
            {
                return UploadValidationResult.Failure(UploadRejectionCode.DocxTooManyEntries,
                    "Archiwum DOCX zawiera zbyt wiele elementów.");
            }

            long totalUncompressed = 0;
            var hasContentTypes = false;
            var hasRootRels = false;

            foreach (var entry in archive.Entries)
            {
                var fullName = entry.FullName.Replace('\\', '/');
                if (IsZipSlipRisk(fullName))
                {
                    return UploadValidationResult.Failure(UploadRejectionCode.DocxZipSlipRisk,
                        "Archiwum DOCX zawiera niedozwoloną ścieżkę.");
                }

                if (IsSuspiciousDocxEntry(fullName))
                {
                    return UploadValidationResult.Failure(UploadRejectionCode.DocxSuspiciousEntry,
                        "Archiwum DOCX zawiera niedozwolony typ pliku.");
                }

                if (_options.RejectDocxMacros && fullName.Equals("word/vbaProject.bin", StringComparison.OrdinalIgnoreCase))
                {
                    return UploadValidationResult.Failure(UploadRejectionCode.DocxSuspiciousEntry,
                        "Archiwum DOCX zawiera makra VBA, które są niedozwolone.");
                }

                if (fullName.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
                    hasContentTypes = true;

                if (fullName.Equals("_rels/.rels", StringComparison.OrdinalIgnoreCase))
                    hasRootRels = true;

                totalUncompressed += entry.Length;
                if (totalUncompressed > _options.MaxDocxUncompressedBytes)
                {
                    return UploadValidationResult.Failure(UploadRejectionCode.DocxTooLargeUncompressed,
                        "Archiwum DOCX po rozpakowaniu przekracza dopuszczalny rozmiar.");
                }

                if (entry.CompressedLength > 0)
                {
                    var ratio = (double)entry.Length / entry.CompressedLength;
                    if (ratio > _options.MaxCompressionRatio)
                    {
                        return UploadValidationResult.Failure(UploadRejectionCode.DocxCompressionRatioExceeded,
                            "Archiwum DOCX ma podejrzanie wysoki współczynnik kompresji.");
                    }
                }
            }

            if (!hasContentTypes || !hasRootRels)
            {
                return UploadValidationResult.Failure(UploadRejectionCode.DocxRequiredPartMissing,
                    "Archiwum DOCX nie zawiera wymaganych elementów OOXML.");
            }

            return UploadValidationResult.Success(DocumentMimeByExtension[".docx"]);
        }
        catch (InvalidDataException)
        {
            return UploadValidationResult.Failure(UploadRejectionCode.DocxInvalidArchive,
                "Plik DOCX jest uszkodzony lub nieprawidłowy.");
        }
    }

    private UploadValidationResult? EvaluateScanFailure(FileScanResult scan)
    {
        return scan.Status switch
        {
            FileScanStatus.Clean => null,
            FileScanStatus.Infected => UploadValidationResult.Failure(UploadRejectionCode.MalwareDetected,
                "Plik został odrzucony przez skaner bezpieczeństwa."),
            FileScanStatus.Unavailable or FileScanStatus.Error when _options.FailClosedWhenScannerFails =>
                UploadValidationResult.Failure(UploadRejectionCode.MalwareScanUnavailable,
                    "Skaner bezpieczeństwa jest chwilowo niedostępny. Spróbuj ponownie później."),
            _ => null
        };
    }

    private static bool IsZipSlipRisk(string entryName) =>
        entryName.StartsWith("/", StringComparison.Ordinal)
        || entryName.StartsWith("../", StringComparison.Ordinal)
        || entryName.Contains("/../", StringComparison.Ordinal)
        || entryName.Contains(':', StringComparison.Ordinal);

    private static bool IsSuspiciousDocxEntry(string entryName)
    {
        var ext = Path.GetExtension(entryName);
        return ext.Length > 0 && SuspiciousDocxExtensions.Contains(ext);
    }

    /// <summary>Nagłówek OLE Compound File (binarny .doc, także zaszyfrowany DOCX): D0 CF 11 E0 A1 B1 1A E1.</summary>
    private static bool LooksLikeCfb(byte[] content) =>
        content.Length >= 8
        && content[0] == 0xD0 && content[1] == 0xCF && content[2] == 0x11 && content[3] == 0xE0
        && content[4] == 0xA1 && content[5] == 0xB1 && content[6] == 0x1A && content[7] == 0xE1;

    private static bool LooksLikePdf(byte[] content) =>
        content.Length >= 5
        && content[0] == 0x25
        && content[1] == 0x50
        && content[2] == 0x44
        && content[3] == 0x46
        && content[4] == 0x2D;

    private static bool LooksLikeZip(byte[] content)
    {
        if (content.Length < 4) return false;

        var b0 = content[0];
        var b1 = content[1];
        var b2 = content[2];
        var b3 = content[3];

        return b0 == 0x50 && b1 == 0x4B
            && ((b2 == 0x03 && b3 == 0x04) || (b2 == 0x05 && b3 == 0x06) || (b2 == 0x07 && b3 == 0x08));
    }

    private static bool MatchesImageSignature(string expectedMime, byte[] content)
    {
        return expectedMime switch
        {
            "image/jpeg" => content.Length >= 3 && content[0] == 0xFF && content[1] == 0xD8 && content[2] == 0xFF,
            "image/png" => content.Length >= 8
                           && content[0] == 0x89 && content[1] == 0x50 && content[2] == 0x4E && content[3] == 0x47
                           && content[4] == 0x0D && content[5] == 0x0A && content[6] == 0x1A && content[7] == 0x0A,
            "image/gif" => content.Length >= 6
                           && content[0] == 0x47 && content[1] == 0x49 && content[2] == 0x46
                           && content[3] == 0x38 && (content[4] == 0x37 || content[4] == 0x39) && content[5] == 0x61,
            "image/bmp" => content.Length >= 2 && content[0] == 0x42 && content[1] == 0x4D,
            "image/webp" => content.Length >= 12
                            && content[0] == 0x52 && content[1] == 0x49 && content[2] == 0x46 && content[3] == 0x46
                            && content[8] == 0x57 && content[9] == 0x45 && content[10] == 0x42 && content[11] == 0x50,
            _ => false
        };
    }
}
