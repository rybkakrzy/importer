namespace Importer.Domain.Models;

/// <summary>
/// Nagłówek lub stopka dokumentu
/// </summary>
public class HeaderFooterContent
{
    public string Html { get; set; } = string.Empty;
    public double Height { get; set; } = 1.25;
    public bool DifferentFirstPage { get; set; }
    public string? FirstPageHtml { get; set; }
}

/// <summary>
/// Reprezentuje dokument z konwersji DOCX do HTML
/// </summary>
public class DocumentContent
{
    public string Html { get; set; } = string.Empty;
    public DocumentMetadata Metadata { get; set; } = new();
    public List<DocumentImage> Images { get; set; } = new();
    public List<DocumentStyle> Styles { get; set; } = new();
    public HeaderFooterContent? Header { get; set; }
    public HeaderFooterContent? Footer { get; set; }
}

/// <summary>
/// Metadane dokumentu DOCX
/// </summary>
public class DocumentMetadata
{
    public string? Title { get; set; }
    public string? Author { get; set; }
    public string? Subject { get; set; }
    public DateTime? Created { get; set; }
    public DateTime? Modified { get; set; }
    public int PageCount { get; set; }
    public int WordCount { get; set; }
}

/// <summary>
/// Obraz osadzony w dokumencie
/// </summary>
public class DocumentImage
{
    public string Id { get; set; } = string.Empty;
    public string ContentType { get; set; } = string.Empty;
    public string Base64Data { get; set; } = string.Empty;
}

/// <summary>
/// Request do zapisu dokumentu
/// </summary>
public class SaveDocumentRequest
{
    public string Html { get; set; } = string.Empty;
    public string? OriginalFileName { get; set; }
    public DocumentMetadata? Metadata { get; set; }
    public HeaderFooterContent? Header { get; set; }
    public HeaderFooterContent? Footer { get; set; }
}

/// <summary>
/// Styl paragrafu
/// </summary>
public class ParagraphStyle
{
    public string? FontFamily { get; set; }
    public double? FontSize { get; set; }
    public string? Color { get; set; }
    public string? BackgroundColor { get; set; }
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public bool IsUnderline { get; set; }
    public bool IsStrikethrough { get; set; }
    public string? Alignment { get; set; }
    public double? LineSpacing { get; set; }
    public double? SpaceBefore { get; set; }
    public double? SpaceAfter { get; set; }
    public double? LeftIndent { get; set; }
    public double? RightIndent { get; set; }
    public double? FirstLineIndent { get; set; }
}

/// <summary>
/// Styl dokumentu (Nagłówek 1, Normalny, itp.)
/// </summary>
public class DocumentStyle
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Type { get; set; } = "paragraph";
    public string? BasedOn { get; set; }
    public string? NextStyle { get; set; }

    public string? FontFamily { get; set; }
    public double? FontSize { get; set; }
    public string? Color { get; set; }
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public bool IsUnderline { get; set; }

    public string? Alignment { get; set; }
    public double? SpaceBefore { get; set; }
    public double? SpaceAfter { get; set; }
    public double? LineSpacing { get; set; }
    public double? LeftIndent { get; set; }
    public double? RightIndent { get; set; }
    public double? FirstLineIndent { get; set; }

    public int? OutlineLevel { get; set; }
}

/// <summary>
/// Domyślne style Word
/// </summary>
public static class DefaultWordStyles
{
    public static List<DocumentStyle> GetDefaultStyles()
    {
        return new List<DocumentStyle>
        {
            new()
            {
                Id = "Normal", Name = "Normalny", Type = "paragraph",
                FontFamily = "Calibri", FontSize = 11, Color = "#000000",
                Alignment = "left", SpaceAfter = 8, LineSpacing = 1.08
            },
            new()
            {
                Id = "Heading1", Name = "Nagłówek 1", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri Light", FontSize = 16, Color = "#2F5496",
                IsBold = true, SpaceBefore = 12, SpaceAfter = 0, OutlineLevel = 1
            },
            new()
            {
                Id = "Heading2", Name = "Nagłówek 2", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri Light", FontSize = 13, Color = "#2F5496",
                IsBold = true, SpaceBefore = 2, SpaceAfter = 0, OutlineLevel = 2
            },
            new()
            {
                Id = "Heading3", Name = "Nagłówek 3", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri Light", FontSize = 12, Color = "#1F3763",
                IsBold = true, SpaceBefore = 2, SpaceAfter = 0, OutlineLevel = 3
            },
            new()
            {
                Id = "Heading4", Name = "Nagłówek 4", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri Light", FontSize = 11, Color = "#2F5496",
                IsBold = true, IsItalic = true, SpaceBefore = 2, SpaceAfter = 0, OutlineLevel = 4
            },
            new()
            {
                Id = "Heading5", Name = "Nagłówek 5", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri Light", FontSize = 11, Color = "#2F5496",
                SpaceBefore = 2, SpaceAfter = 0, OutlineLevel = 5
            },
            new()
            {
                Id = "Heading6", Name = "Nagłówek 6", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri Light", FontSize = 11, Color = "#1F3763",
                IsItalic = true, SpaceBefore = 2, SpaceAfter = 0, OutlineLevel = 6
            },
            new()
            {
                Id = "Title", Name = "Tytuł", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri Light", FontSize = 28, Color = "#000000",
                SpaceAfter = 0, LineSpacing = 1.0
            },
            new()
            {
                Id = "Subtitle", Name = "Podtytuł", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri", FontSize = 11, Color = "#5A5A5A",
                IsItalic = true, SpaceAfter = 8
            },
            new()
            {
                Id = "Quote", Name = "Cytat", Type = "paragraph",
                BasedOn = "Normal", NextStyle = "Normal",
                FontFamily = "Calibri", FontSize = 11, Color = "#404040",
                IsItalic = true, LeftIndent = 1.27, RightIndent = 1.27,
                SpaceBefore = 10, SpaceAfter = 10
            },
            new()
            {
                Id = "ListParagraph", Name = "Akapit listy", Type = "paragraph",
                BasedOn = "Normal", LeftIndent = 1.27
            }
        };
    }
}
