namespace D2ViewerEditor.Infrastructure.DocxModel;

public enum PageOrientation
{
    Portrait,
    Landscape
}

/// <summary>
/// First piece of the explicit intermediate document model (Etap 2). Holds a section's
/// page geometry as raw OOXML twips exactly as authored — callers map to cm/px through
/// <c>OoxmlUnits</c>. Parsing is separated from rendering so section/page handling is
/// testable on the model instead of only on the final HTML string.
///
/// Page size and orientation are parsed now (even though the current HTML output does
/// not yet surface them) to seed Etap 4 without a second parse pass.
/// </summary>
public sealed class PageSettings
{
    public int? PageWidthTwips { get; init; }
    public int? PageHeightTwips { get; init; }
    public PageOrientation Orientation { get; init; } = PageOrientation.Portrait;

    /// <summary>True when the section declares a w:pgMar element at all.</summary>
    public bool HasPageMargin { get; init; }

    public int? TopMarginTwips { get; init; }
    public int? BottomMarginTwips { get; init; }
    public int? LeftMarginTwips { get; init; }
    public int? RightMarginTwips { get; init; }
    public int? HeaderDistanceTwips { get; init; }
    public int? FooterDistanceTwips { get; init; }
}
