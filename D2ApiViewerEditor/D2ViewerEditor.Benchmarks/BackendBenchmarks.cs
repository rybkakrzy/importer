using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Services;

namespace D2ViewerEditor.Benchmarks;

internal static class Pipeline
{
    public static DocumentContent Read(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        return new DocxToHtmlConverter().Convert(ms);
    }

    public static byte[] Write(DocumentContent content) =>
        new HtmlToDocxConverter().Convert(content.Html, content.Metadata, content.Header,
            content.Footer, content.Margins, content.PageSize);

    public static byte[] WritePreserving(DocumentContent content, byte[] original)
    {
        using var ms = new MemoryStream(original);
        return new HtmlToDocxConverter().ConvertPreservingPackage(content.Html, ms, content.Metadata,
            content.Header, content.Footer, content.Margins, content.PageSize);
    }
}

/// <summary>
/// DOCX → DocumentContent (reader). Measures parse + DocumentContent build time and allocations
/// for representative document shapes. Returns Html length to defeat dead-code elimination.
/// </summary>
[MemoryDiagnoser]
[MarkdownExporterAttribute.GitHub]
public class DocxImportBenchmarks
{
    private byte[] _simple = null!, _tables = null!, _images = null!, _multiStyle = null!;
    private byte[]? _regression;

    [GlobalSetup]
    public void Setup()
    {
        _simple = BenchmarkAssets.Simple();
        _tables = BenchmarkAssets.Tables();
        _images = BenchmarkAssets.Images();
        _multiStyle = BenchmarkAssets.MultiStyle();
        _regression = BenchmarkAssets.TryLoadRegressionOriginal();
    }

    [Benchmark(Baseline = true)] public int Simple() => Pipeline.Read(_simple).Html.Length;
    [Benchmark] public int Tables() => Pipeline.Read(_tables).Html.Length;
    [Benchmark] public int Images() => Pipeline.Read(_images).Html.Length;
    [Benchmark] public int MultiStyle() => Pipeline.Read(_multiStyle).Html.Length;
    [Benchmark] public int Regression() => _regression == null ? 0 : Pipeline.Read(_regression).Html.Length;
}

/// <summary>
/// DocumentContent/HTML → DOCX (writer). Covers the self-contained writer and the pass-through
/// writer (preserving the original styles/theme/fontTable — R-16). Returns byte count.
/// </summary>
[MemoryDiagnoser]
[MarkdownExporterAttribute.GitHub]
public class DocxExportBenchmarks
{
    private DocumentContent _simple = null!, _tables = null!;
    private DocumentContent? _regression;
    private byte[] _simpleSrc = null!, _tablesSrc = null!;
    private byte[]? _regressionSrc;

    [GlobalSetup]
    public void Setup()
    {
        _simpleSrc = BenchmarkAssets.Simple();
        _tablesSrc = BenchmarkAssets.Tables();
        _regressionSrc = BenchmarkAssets.TryLoadRegressionOriginal();
        _simple = Pipeline.Read(_simpleSrc);
        _tables = Pipeline.Read(_tablesSrc);
        _regression = _regressionSrc == null ? null : Pipeline.Read(_regressionSrc);
    }

    [Benchmark(Baseline = true)] public int ExportSimple() => Pipeline.Write(_simple).Length;
    [Benchmark] public int ExportTables() => Pipeline.Write(_tables).Length;
    [Benchmark] public int ExportPassThroughTables() => Pipeline.WritePreserving(_tables, _tablesSrc).Length;
    [Benchmark] public int ExportRegression() => _regression == null ? 0 : Pipeline.Write(_regression).Length;
    [Benchmark] public int ExportPassThroughRegression() =>
        _regression == null || _regressionSrc == null ? 0 : Pipeline.WritePreserving(_regression, _regressionSrc).Length;
}

/// <summary>
/// Full DOCX → DocumentContent → HTML → DOCX round-trip. The headline pipeline metric: any
/// regression in reader or writer shows up here in time and allocations.
/// </summary>
[MemoryDiagnoser]
[MarkdownExporterAttribute.GitHub]
public class DocxRoundTripBenchmarks
{
    private byte[] _tables = null!, _headerFooter = null!, _pageBreaks = null!;
    private byte[]? _regression;

    [GlobalSetup]
    public void Setup()
    {
        _tables = BenchmarkAssets.Tables();
        _headerFooter = BenchmarkAssets.HeaderFooter();
        _pageBreaks = BenchmarkAssets.PageBreaks();
        _regression = BenchmarkAssets.TryLoadRegressionOriginal();
    }

    [Benchmark(Baseline = true)] public int Tables() => Pipeline.Write(Pipeline.Read(_tables)).Length;
    [Benchmark] public int HeaderFooter() => Pipeline.Write(Pipeline.Read(_headerFooter)).Length;
    [Benchmark] public int PageBreaks() => Pipeline.Write(Pipeline.Read(_pageBreaks)).Length;

    [Benchmark]
    public int RegressionPassThrough()
    {
        if (_regression == null) return 0;
        var content = Pipeline.Read(_regression);
        return Pipeline.WritePreserving(content, _regression).Length;
    }
}

/// <summary>
/// Table-heavy paths (parse / serialize large tables) — the area most sensitive to per-cell
/// work. ShortRun keeps the large-table cases quick enough for routine local runs.
/// </summary>
[MemoryDiagnoser]
[ShortRunJob]
[MarkdownExporterAttribute.GitHub]
public class TableBenchmarks
{
    [Params(100, 400)] public int Rows;

    private byte[] _src = null!;
    private DocumentContent _content = null!;

    [GlobalSetup]
    public void Setup()
    {
        _src = BenchmarkAssets.LargeTable(Rows);
        _content = Pipeline.Read(_src);
    }

    [Benchmark] public int ParseLargeTable() => Pipeline.Read(_src).Html.Length;
    [Benchmark] public int SerializeLargeTable() => Pipeline.Write(_content).Length;
}
