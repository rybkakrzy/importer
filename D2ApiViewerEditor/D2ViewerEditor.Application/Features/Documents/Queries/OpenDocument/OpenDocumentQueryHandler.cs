using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Domain.Models;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.OpenDocument;

public class OpenDocumentQueryHandler : IRequestHandler<OpenDocumentQuery, Result<DocumentContent>>
{
    // Sentinele rozpoznawane przez kontroler → mapowane na właściwe kody HTTP / payload dla GUI.
    public const string PasswordRequiredSentinel = "PASSWORD_REQUIRED";
    public const string WrongPasswordSentinel = "WRONG_PASSWORD";
    public const string UnsupportedLegacyDocSentinel = "UNSUPPORTED_LEGACY_DOC";

    private readonly IDocxToHtmlConverter _converter;
    private readonly IDocumentInputNormalizer _normalizer;

    public OpenDocumentQueryHandler(IDocxToHtmlConverter converter, IDocumentInputNormalizer normalizer)
    {
        _converter = converter;
        _normalizer = normalizer;
    }

    public async Task<Result<DocumentContent>> Handle(OpenDocumentQuery request, CancellationToken cancellationToken)
    {
        using var memoryStream = new MemoryStream();
        await request.FileStream.CopyToAsync(memoryStream, cancellationToken);

        // Doprowadź wejście do zwykłego DOCX: pass-through ZIP, dekrypcja DOCX z hasłem, detekcja .doc.
        var normalized = _normalizer.Normalize(memoryStream.ToArray(), request.Password);
        switch (normalized.Status)
        {
            case DocumentInputStatus.Ok:
                break;
            case DocumentInputStatus.PasswordRequired:
                return Result<DocumentContent>.Failure(PasswordRequiredSentinel);
            case DocumentInputStatus.WrongPassword:
                return Result<DocumentContent>.Failure(WrongPasswordSentinel);
            case DocumentInputStatus.UnsupportedLegacyDoc:
                return Result<DocumentContent>.Failure(UnsupportedLegacyDocSentinel);
            default:
                return Result<DocumentContent>.Failure("Nie rozpoznano formatu pliku lub plik jest uszkodzony.");
        }

        using var docxStream = new MemoryStream(normalized.Docx!);
        var content = _converter.Convert(docxStream);
        content.Metadata.Title ??= Path.GetFileNameWithoutExtension(request.FileName);

        return Result<DocumentContent>.Success(content);
    }
}
