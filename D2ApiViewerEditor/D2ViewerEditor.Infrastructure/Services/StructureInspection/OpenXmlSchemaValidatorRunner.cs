using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using D2ViewerEditor.Domain.Models;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Infrastructure.Services.StructureInspection;

/// <summary>Wynik walidacji schematu: lista przycięta do limitu prezentacyjnego oraz pełny licznik.</summary>
public sealed record SchemaValidationResult(IReadOnlyList<SchemaValidationIssue> Issues, int TotalCount);

/// <summary>
/// Walidacja pakietu przez <c>OpenXmlValidator</c> z wybieralnym profilem Office. To warstwa
/// interpretacji SDK, niezależna od surowego XML: odpowiada na pytanie „czy pakiet jest zgodny ze
/// schematem", a nie „czy Word i nasz edytor odtworzą dokument tak samo".
/// </summary>
public sealed class OpenXmlSchemaValidatorRunner
{
    private readonly StructureInspectionOptions _options;

    public OpenXmlSchemaValidatorRunner(IOptions<StructureInspectionOptions> options)
    {
        _options = options.Value;
    }

    public static IReadOnlyList<string> SupportedTargetVersions() => Enum.GetNames<FileFormatVersions>();

    public SchemaValidationResult Validate(byte[] documentBytes, string targetVersion, CancellationToken cancellationToken)
    {
        if (!Enum.TryParse<FileFormatVersions>(targetVersion, ignoreCase: true, out var version))
        {
            throw new ArgumentException($"Nieznany profil Open XML '{targetVersion}'.", nameof(targetVersion));
        }

        using var stream = new MemoryStream(documentBytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);

        var validator = new OpenXmlValidator(version);
        var issues = new List<SchemaValidationIssue>();
        var totalCount = 0;

        foreach (var error in validator.Validate(document))
        {
            cancellationToken.ThrowIfCancellationRequested();
            totalCount++;

            if (issues.Count < _options.MaxSchemaIssues)
            {
                issues.Add(Map(error, targetVersion));
            }
        }

        return new SchemaValidationResult(issues, totalCount);
    }

    private static SchemaValidationIssue Map(ValidationErrorInfo error, string targetVersion) => new(
        string.IsNullOrEmpty(error.Id) ? "OPENXML_VALIDATION" : error.Id,
        error.ErrorType.ToString(),
        error.Description ?? "Błąd walidacji Open XML.",
        error.Part?.Uri.ToString().TrimStart('/'),
        error.Node?.LocalName,
        error.Path?.XPath,
        null,
        targetVersion);
}
