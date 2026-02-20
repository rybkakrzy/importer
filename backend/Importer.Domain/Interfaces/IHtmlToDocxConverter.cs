using Importer.Domain.Models;

namespace Importer.Domain.Interfaces;

/// <summary>
/// Kontrakt konwertera HTML → DOCX
/// </summary>
public interface IHtmlToDocxConverter
{
    /// <summary>
    /// Konwertuje HTML na dokument DOCX (tablica bajtów)
    /// </summary>
    byte[] Convert(string html, DocumentMetadata? metadata = null,
                   HeaderFooterContent? header = null, HeaderFooterContent? footer = null);
}
