using D2Tools.Domain.Models;

namespace D2Tools.Domain.Interfaces;

/// <summary>
/// Kontrakt konwertera DOCX → HTML
/// </summary>
public interface IDocxToHtmlConverter
{
    /// <summary>
    /// Konwertuje strumień DOCX na strukturę DocumentContent (HTML + metadane + style)
    /// </summary>
    DocumentContent Convert(Stream docxStream);
}
