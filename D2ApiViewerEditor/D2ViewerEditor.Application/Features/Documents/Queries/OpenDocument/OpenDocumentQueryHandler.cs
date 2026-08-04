using D2ViewerEditor.Application.Common;
using D2ViewerEditor.Application.Common.Security;
using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Domain.Models;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.OpenDocument;

public class OpenDocumentQueryHandler : IRequestHandler<OpenDocumentQuery, Result<DocumentContent>>
{
    private readonly IDocxToHtmlConverter _converter;
    private readonly IDocumentInputNormalizer _normalizer;
    private readonly IFileUploadSecurityService _uploadSecurity;

    public OpenDocumentQueryHandler(
        IDocxToHtmlConverter converter,
        IDocumentInputNormalizer normalizer,
        IFileUploadSecurityService uploadSecurity)
    {
        _converter = converter;
        _normalizer = normalizer;
        _uploadSecurity = uploadSecurity;
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
                return Result<DocumentContent>.Failure(ErrorCodes.DocumentProtected);
            case DocumentInputStatus.WrongPassword:
                return Result<DocumentContent>.Failure(ErrorCodes.DocumentUnlockFailed);
            case DocumentInputStatus.UnsupportedLegacyDoc:
                return Result<DocumentContent>.Failure(ErrorCodes.UnsupportedLegacyDoc);
            default:
                return Result<DocumentContent>.Failure(ErrorCodes.DocumentFormatInvalid);
        }

        // Bug 13625398: ta sama walidacja struktury co upload ze strony startowej — bez niej
        // uszkodzony DOCX (np. bez [Content_Types].xml) otwierał się „po cichu" jako pusty
        // dokument i komunikaty obu ścieżek były niespójne.
        var structure = _uploadSecurity.ValidateDocxStructure(normalized.Docx!);
        if (!structure.IsValid)
        {
            return Result<DocumentContent>.Failure(structure.Code == UploadRejectionCode.SignatureMismatch
                ? ErrorCodes.DocumentFormatInvalid
                : ErrorCodes.DocumentCorrupted);
        }

        using var docxStream = new MemoryStream(normalized.Docx!);
        var content = _converter.Convert(docxStream);
        content.Metadata.Title ??= Path.GetFileNameWithoutExtension(request.FileName);

        return Result<DocumentContent>.Success(content);
    }
}
