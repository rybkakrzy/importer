using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Domain.Models;
using OoxmlPageSize = DocumentFormat.OpenXml.Wordprocessing.PageSize;

namespace D2ViewerEditor.Infrastructure.DocxModel;

/// <summary>
/// Parses a w:sectPr into the <see cref="PageSettings"/> model. Pure read-side mapping:
/// no defaulting, no unit conversion — values are the raw twips as authored, with nulls
/// preserved so callers keep their existing fallback semantics.
/// </summary>
public static class SectionPropertiesReader
{
    public static PageSettings ReadPageSettings(SectionProperties? sectPr)
    {
        var pageSize = sectPr?.GetFirstChild<OoxmlPageSize>();
        var pageMargin = sectPr?.GetFirstChild<PageMargin>();
        var docGrid = sectPr?.GetFirstChild<DocGrid>();

        var orientation = pageSize?.Orient?.Value == PageOrientationValues.Landscape
            ? PageOrientation.Landscape
            : PageOrientation.Portrait;

        return new PageSettings
        {
            PageWidthTwips = pageSize?.Width?.Value is { } w ? (int)w : null,
            PageHeightTwips = pageSize?.Height?.Value is { } h ? (int)h : null,
            Orientation = orientation,
            HasPageMargin = pageMargin != null,
            TopMarginTwips = pageMargin?.Top?.Value,
            BottomMarginTwips = pageMargin?.Bottom?.Value,
            LeftMarginTwips = pageMargin?.Left?.Value is { } l ? (int)l : null,
            RightMarginTwips = pageMargin?.Right?.Value is { } r ? (int)r : null,
            HeaderDistanceTwips = pageMargin?.Header?.Value is { } hd ? (int)hd : null,
            FooterDistanceTwips = pageMargin?.Footer?.Value is { } fd ? (int)fd : null,
            Columns = ReadColumns(sectPr?.GetFirstChild<Columns>()),
            DocGridType = docGrid == null ? null : DocGridTypeName(docGrid.Type?.Value),
            DocGridLinePitchTwips = docGrid?.LinePitch?.Value,
            DocGridCharSpace = docGrid?.CharacterSpace?.Value,
        };
    }

    /// <summary>w:docGrid/@w:type jako token OOXML (brak atrybutu = "default").</summary>
    private static string DocGridTypeName(DocGridValues? type)
    {
        if (type == DocGridValues.Lines) return "lines";
        if (type == DocGridValues.LinesAndChars) return "linesAndChars";
        if (type == DocGridValues.SnapToChars) return "snapToChars";
        return "default";
    }

    /// <summary>
    /// Maps w:cols to <see cref="ColumnLayout"/>. Null when the section declares no w:cols.
    /// w:num defaults to 1, w:space to 720 twips, w:equalWidth to true (OOXML defaults). When
    /// w:equalWidth is false and explicit w:col children exist, their w:w / w:space are captured
    /// for exact round-trip; otherwise only Count + SpaceTwips are needed (equal columns).
    /// </summary>
    private static ColumnLayout? ReadColumns(Columns? cols)
    {
        if (cols == null) return null;

        // w:equalWidth default is true; only "0"/"false" means unequal.
        var equalWidth = cols.EqualWidth?.Value ?? true;

        var colChildren = cols.Elements<Column>().ToList();

        // w:num is the authoritative count; when absent, fall back to the number of w:col
        // children (unequal layout), else 1.
        int count = cols.ColumnCount?.Value
            ?? (colChildren.Count > 0 ? colChildren.Count : 1);

        var layout = new ColumnLayout
        {
            Count = count < 1 ? 1 : count,
            EqualWidth = equalWidth,
            SpaceTwips = cols.Space?.Value is { } sp && int.TryParse(sp, out var spi) ? spi : 720,
            Separator = cols.Separator?.Value ?? false,
        };

        if (!equalWidth && colChildren.Count > 0)
        {
            layout.Columns = colChildren.Select(c => new SectionColumn
            {
                WidthTwips = c.Width?.Value is { } w && int.TryParse(w, out var wi) ? wi : 0,
                SpaceTwips = c.Space?.Value is { } s && int.TryParse(s, out var si) ? si : 0,
            }).ToList();
        }

        return layout;
    }
}
