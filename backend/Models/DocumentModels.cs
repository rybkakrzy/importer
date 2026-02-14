namespace ImporterApi.Models;

/// <summary>
/// Reprezentuje dokument z konwersji DOCX do HTML
/// </summary>
public class DocumentContent
{
    public string Html { get; set; } = string.Empty;
    public DocumentMetadata Metadata { get; set; } = new();
    public List<DocumentImage> Images { get; set; } = new();
    public List<DocumentStyle> Styles { get; set; } = new();
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
    public string Type { get; set; } = "paragraph"; // paragraph, character
    public string? BasedOn { get; set; }
    public string? NextStyle { get; set; }
    
    // Właściwości czcionki
    public string? FontFamily { get; set; }
    public double? FontSize { get; set; } // w punktach
    public string? Color { get; set; }
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public bool IsUnderline { get; set; }
    
    // Właściwości paragrafu
    public string? Alignment { get; set; } // left, center, right, justify
    public double? SpaceBefore { get; set; } // w punktach
    public double? SpaceAfter { get; set; } // w punktach
    public double? LineSpacing { get; set; } // mnożnik (1.0, 1.5, 2.0)
    public double? LeftIndent { get; set; } // w cm
    public double? RightIndent { get; set; } // w cm
    public double? FirstLineIndent { get; set; } // w cm
    
    // Poziom outline (dla nagłówków)
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
            new DocumentStyle
            {
                Id = "Normal",
                Name = "Normalny",
                Type = "paragraph",
                FontFamily = "Calibri",
                FontSize = 11,
                Color = "#000000",
                Alignment = "left",
                SpaceAfter = 8,
                LineSpacing = 1.08
            },
            new DocumentStyle
            {
                Id = "Heading1",
                Name = "Nagłówek 1",
                Type = "paragraph",
                BasedOn = "Normal",
                NextStyle = "Normal",
                FontFamily = "Calibri Light",
                FontSize = 16,
                Color = "#2F5496",
                IsBold = true,
                SpaceBefore = 12,
                SpaceAfter = 0,
                OutlineLevel = 1
            },
            new DocumentStyle
            {
                Id = "Heading2",
                Name = "Nagłówek 2",
                Type = "paragraph",
                BasedOn = "Normal",
                NextStyle = "Normal",
                FontFamily = "Calibri Light",
                FontSize = 13,
                Color = "#2F5496",
                IsBold = true,
                SpaceBefore = 2,
                SpaceAfter = 0,
                OutlineLevel = 2
            },
            new DocumentStyle
            {
                Id = "Heading3",
                Name = "Nagłówek 3",
                Type = "paragraph",
                BasedOn = "Normal",
                NextStyle = "Normal",
                FontFamily = "Calibri Light",
                FontSize = 12,
                Color = "#1F3763",
                IsBold = true,
                SpaceBefore = 2,
                SpaceAfter = 0,
                OutlineLevel = 3
            },
            new DocumentStyle
            {
                Id = "Heading4",
                Name = "Nagłówek 4",
                Type = "paragraph",
                BasedOn = "Normal",
                NextStyle = "Normal",
                FontFamily = "Calibri Light",
                FontSize = 11,
                Color = "#2F5496",
                IsBold = true,
                IsItalic = true,
                SpaceBefore = 2,
                SpaceAfter = 0,
                OutlineLevel = 4
            },
            new DocumentStyle
            {
                Id = "Heading5",
                Name = "Nagłówek 5",
                Type = "paragraph",
                BasedOn = "Normal",
                NextStyle = "Normal",
                FontFamily = "Calibri Light",
                FontSize = 11,
                Color = "#2F5496",
                SpaceBefore = 2,
                SpaceAfter = 0,
                OutlineLevel = 5
            },
            new DocumentStyle
            {
                Id = "Heading6",
                Name = "Nagłówek 6",
                Type = "paragraph",
                BasedOn = "Normal",
                NextStyle = "Normal",
                FontFamily = "Calibri Light",
                FontSize = 11,
                Color = "#1F3763",
                IsItalic = true,
                SpaceBefore = 2,
                SpaceAfter = 0,
                OutlineLevel = 6
            },
            new DocumentStyle
            {
                Id = "Title",
                Name = "Tytuł",
                Type = "paragraph",
                BasedOn = "Normal",
                NextStyle = "Normal",
                FontFamily = "Calibri Light",
                FontSize = 28,
                Color = "#000000",
                SpaceAfter = 0,
                LineSpacing = 1.0
            },
            new DocumentStyle
            {
                Id = "Subtitle",
                Name = "Podtytuł",
                Type = "paragraph",
                BasedOn = "Normal",
                NextStyle = "Normal",
                FontFamily = "Calibri",
                FontSize = 11,
                Color = "#5A5A5A",
                IsItalic = true,
                SpaceAfter = 8
            },
            new DocumentStyle
            {
                Id = "Quote",
                Name = "Cytat",
                Type = "paragraph",
                BasedOn = "Normal",
                NextStyle = "Normal",
                FontFamily = "Calibri",
                FontSize = 11,
                Color = "#404040",
                IsItalic = true,
                LeftIndent = 1.27,
                RightIndent = 1.27,
                SpaceBefore = 10,
                SpaceAfter = 10
            },
            new DocumentStyle
            {
                Id = "ListParagraph",
                Name = "Akapit listy",
                Type = "paragraph",
                BasedOn = "Normal",
                LeftIndent = 1.27
            }
        };
    }
}
