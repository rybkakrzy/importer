using D2ViewerEditor.Application.Features.Documents.Common;
using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;
using Microsoft.Extensions.Logging;

namespace D2ViewerEditor.Application.Features.Documents.Commands.DownloadEditedDocument;

public class DownloadEditedDocumentCommandHandler
    : IRequestHandler<DownloadEditedDocumentCommand, Result<DownloadEditedDocumentResult>>
{
    /// <summary>
    /// Sentinel prefix the controller maps to HTTP 403. Distinct from "not found" /
    /// generic failure so we never silently leak bytes when the gate trips.
    /// </summary>
    public const string ForbiddenErrorPrefix = "USER_DOWNLOAD_FORBIDDEN:";

    private readonly IDocumentRepository _documentRepository;
    private readonly IHtmlToDocxConverter _converter;
    private readonly ILogger<DownloadEditedDocumentCommandHandler> _logger;

    public DownloadEditedDocumentCommandHandler(
        IDocumentRepository documentRepository,
        IHtmlToDocxConverter converter,
        ILogger<DownloadEditedDocumentCommandHandler> logger)
    {
        _documentRepository = documentRepository;
        _converter = converter;
        _logger = logger;
    }

    public async Task<Result<DownloadEditedDocumentResult>> Handle(
        DownloadEditedDocumentCommand request,
        CancellationToken cancellationToken)
    {
        var document = await _documentRepository.GetByIdAsync(request.MasterId, cancellationToken);
        if (document == null)
            return Result<DownloadEditedDocumentResult>.NotFound();

        var meta = ExternalDocumentMetadata.Parse(document.Metadata);
        if (!meta.IsUserDownloadAllowed)
        {
            // Log the decision (master only — no editor state, no file content).
            _logger.LogWarning(
                "User-download blocked for document {MasterId}: userDownload flag not set to true.",
                document.Id);
            return Result<DownloadEditedDocumentResult>.Failure(
                ForbiddenErrorPrefix + " Pobieranie pliku na komputer nie jest dostępne dla tego dokumentu.");
        }

        if (string.IsNullOrWhiteSpace(request.Html))
            return Result<DownloadEditedDocumentResult>.Failure("HTML edytora nie może być pusty.");

        var docxBytes = _converter.Convert(
            request.Html, request.Metadata, request.Header, request.Footer, request.Margins);

        var fileName = string.IsNullOrWhiteSpace(request.OriginalFileName)
            ? $"{(string.IsNullOrWhiteSpace(document.Name) ? "dokument" : document.Name)}.docx"
            : request.OriginalFileName;

        return Result<DownloadEditedDocumentResult>.Success(
            new DownloadEditedDocumentResult(docxBytes, fileName));
    }
}
