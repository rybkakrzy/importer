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

    private const string DocxMimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private readonly IDocumentRepository _documentRepository;
    private readonly IHtmlToDocxConverter _converter;
    private readonly IDocumentStorageService _storageService;
    private readonly ILogger<DownloadEditedDocumentCommandHandler> _logger;

    public DownloadEditedDocumentCommandHandler(
        IDocumentRepository documentRepository,
        IHtmlToDocxConverter converter,
        IDocumentStorageService storageService,
        ILogger<DownloadEditedDocumentCommandHandler> logger)
    {
        _documentRepository = documentRepository;
        _converter = converter;
        _storageService = storageService;
        _logger = logger;
    }

    public async Task<Result<DownloadEditedDocumentResult>> Handle(
        DownloadEditedDocumentCommand request,
        CancellationToken cancellationToken)
    {
        var document = await _documentRepository.GetByIdWithVersionsAsync(request.MasterId, cancellationToken);
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

        // Pass-through: when the original DOCX package is available, preserve its styles/theme/
        // fontTable so the edited download keeps the full style set and document fonts (R-16).
        // Falls back to a self-contained conversion when there is no original (e.g. non-DOCX,
        // or no stored version).
        var baseVersion = document.Versions
            .OrderBy(v => v.VersionNumber)
            .FirstOrDefault();

        byte[] docxBytes;
        if (baseVersion != null && document.MimeType == DocxMimeType)
        {
            var original = await _storageService.DownloadAsync(baseVersion.StoragePath, cancellationToken);
            using var originalStream = new MemoryStream(original);
            docxBytes = _converter.ConvertPreservingPackage(
                request.Html, originalStream, request.Metadata, request.Header, request.Footer, request.Margins, request.PageSize,
                request.SectionHeadersFooters, request.Footnotes);
        }
        else
        {
            docxBytes = _converter.Convert(
                request.Html, request.Metadata, request.Header, request.Footer, request.Margins, request.PageSize,
                request.SectionHeadersFooters, request.Footnotes);
        }

        var fileName = string.IsNullOrWhiteSpace(request.OriginalFileName)
            ? $"{(string.IsNullOrWhiteSpace(document.Name) ? "dokument" : document.Name)}.docx"
            : request.OriginalFileName;

        return Result<DownloadEditedDocumentResult>.Success(
            new DownloadEditedDocumentResult(docxBytes, fileName));
    }
}
