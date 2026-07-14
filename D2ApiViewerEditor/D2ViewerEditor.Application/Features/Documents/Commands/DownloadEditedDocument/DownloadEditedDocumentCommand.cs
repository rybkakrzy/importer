using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Models;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.DownloadEditedDocument;

/// <summary>
/// User-facing "download edited document" command. Converts the current editor state
/// (HTML + header/footer/margins/metadata) to DOCX bytes, but only when the document's
/// metadata has <c>userDownload == true</c>. Used by the dedicated gated endpoint on
/// <see cref="DocumentStorageController"/>; the stateless converter on
/// <c>POST /api/document/save</c> is intentionally NOT used for this path so the gate
/// cannot be bypassed by sending a request without <c>masterId</c>.
/// </summary>
public record DownloadEditedDocumentCommand(
    Guid MasterId,
    string Html,
    string? OriginalFileName,
    DocumentMetadata? Metadata,
    HeaderFooterContent? Header,
    HeaderFooterContent? Footer,
    PageMargins? Margins,
    PageSize? PageSize = null,
    List<SectionHeaderFooter>? SectionHeadersFooters = null,
    List<Footnote>? Footnotes = null,
    List<Endnote>? Endnotes = null
) : IRequest<Result<DownloadEditedDocumentResult>>;

public record DownloadEditedDocumentResult(byte[] DocxBytes, string FileName);
