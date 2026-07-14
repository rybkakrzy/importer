using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;
using Microsoft.Extensions.Logging;

namespace D2ViewerEditor.Application.Features.Documents.Commands.SaveDocument;

public class SaveDocumentCommandHandler : IRequestHandler<SaveDocumentCommand, Result<SaveDocumentResult>>
{
    private const string DocxMimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private readonly IHtmlToDocxConverter _converter;
    // Opcjonalne: potrzebne tylko dla pass-through (gdy komenda niesie MasterId). Nullable, by
    // czysta konwersja HTML→DOCX (bez MasterId, np. testy/„nowy dokument") działała bez storage.
    private readonly IDocumentRepository? _documentRepository;
    private readonly IDocumentStorageService? _storageService;
    private readonly ILogger<SaveDocumentCommandHandler>? _logger;

    public SaveDocumentCommandHandler(
        IHtmlToDocxConverter converter,
        IDocumentRepository? documentRepository = null,
        IDocumentStorageService? storageService = null,
        ILogger<SaveDocumentCommandHandler>? logger = null)
    {
        _converter = converter;
        _documentRepository = documentRepository;
        _storageService = storageService;
        _logger = logger;
    }

    public async Task<Result<SaveDocumentResult>> Handle(SaveDocumentCommand request, CancellationToken cancellationToken)
    {
        var docxBytes = await ConvertAsync(request, cancellationToken);

        var fileName = !string.IsNullOrEmpty(request.OriginalFileName)
            ? request.OriginalFileName
            : "dokument.docx";

        if (!fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            fileName += ".docx";

        return Result<SaveDocumentResult>.Success(new SaveDocumentResult(docxBytes, fileName));
    }

    /// <summary>
    /// Pass-through zachowuje styles.xml/theme/fontTable/numbering oryginalnego pakietu, więc
    /// definicje stylów tabel (i cały zestaw ~160 stylów) przeżywają zapis — tabele nie tracą
    /// obramowań/formatowania ze stylu przy autosave (R-16/R-19). Uruchamiany tylko gdy komenda
    /// niesie MasterId, dokument jest DOCX i ma wersję bazową; każdy błąd → best-effort fallback
    /// do pełnej regeneracji (zapis nigdy nie może się wywalić przez pass-through).
    /// </summary>
    private async Task<byte[]> ConvertAsync(SaveDocumentCommand request, CancellationToken cancellationToken)
    {
        if (request.MasterId is { } masterId && _documentRepository != null && _storageService != null)
        {
            try
            {
                var document = await _documentRepository.GetByIdWithVersionsAsync(masterId, cancellationToken);
                if (document != null && document.MimeType == DocxMimeType)
                {
                    var baseVersion = document.Versions.OrderBy(v => v.VersionNumber).FirstOrDefault();
                    if (baseVersion != null)
                    {
                        var original = await _storageService.DownloadAsync(baseVersion.StoragePath, cancellationToken);
                        using var originalStream = new MemoryStream(original);
                        return _converter.ConvertPreservingPackage(
                            request.Html, originalStream, request.Metadata, request.Header, request.Footer,
                            request.Margins, request.PageSize, request.SectionHeadersFooters, request.Footnotes, request.Endnotes);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex,
                    "Pass-through zapisu nie powiódł się dla dokumentu {MasterId} — fallback do pełnej regeneracji pakietu.",
                    masterId);
            }
        }

        return _converter.Convert(
            request.Html, request.Metadata, request.Header, request.Footer,
            request.Margins, request.PageSize, request.SectionHeadersFooters, request.Footnotes, request.Endnotes);
    }
}
