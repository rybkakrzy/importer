using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Models;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.SaveDocument;

/// <summary>
/// Komenda zapisania dokumentu HTML jako DOCX
/// </summary>
public record SaveDocumentCommand(
    string Html,
    string? OriginalFileName,
    DocumentMetadata? Metadata,
    HeaderFooterContent? Header,
    HeaderFooterContent? Footer,
    PageMargins? Margins = null,
    PageSize? PageSize = null,
    List<SectionHeaderFooter>? SectionHeadersFooters = null,
    List<Footnote>? Footnotes = null,
    // Gdy podane i dokument to DOCX z wersją bazową — konwersja idzie przez
    // ConvertPreservingPackage (zachowuje styles.xml/theme/fontTable/numbering oryginału),
    // więc definicje stylów tabel/motywu przeżywają zapis. Brak → pełna regeneracja (jak dotąd).
    Guid? MasterId = null
) : IRequest<Result<SaveDocumentResult>>;

public record SaveDocumentResult(byte[] DocxBytes, string FileName);
