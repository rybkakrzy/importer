using System.Buffers.Binary;
using System.Text;
using BenchmarkDotNet.Attributes;
using D2ViewerEditor.Domain.Models;
using D2ViewerEditor.Infrastructure.Services;

namespace D2ViewerEditor.Benchmarks;

/// <summary>
/// Wydajność/pamięć konwertera grafik (pure-managed, bez LibreOffice/GDI). Mierzy detekcję,
/// passthrough rastra, placeholder EMF, ekstrakcję osadzonego rastra, VML→SVG, sanitizację SVG
/// oraz przepustowość dokumentu z N grafikami. Uruchom: `dotnet run -c Release -- --filter *GraphicConversion*`.
/// Patrz .ai/GRAPHICS_CONVERSION.md (§benchmarki).
/// </summary>
[MemoryDiagnoser]
[MarkdownExporterAttribute.GitHub]
public class GraphicConversionBenchmarks
{
    private readonly GraphicConversionService _svc = new();
    private byte[] _png = null!;
    private byte[] _emf = null!;
    private byte[] _emfWithPng = null!;
    private string _vml = null!;
    private string _svg = null!;

    /// <summary>Liczba grafik w syntetycznym dokumencie (throughput).</summary>
    [Params(1, 10, 100)]
    public int GraphicCount;

    [GlobalSetup]
    public void Setup()
    {
        _png = MinimalPng(64, 48);
        _emf = BuildEmf(10000, 5000);
        _emfWithPng = _emf.Concat(MinimalPng(40, 40)).ToArray();
        _vml = "<v:rect xmlns:v='urn:schemas-microsoft-com:vml' style='width:100pt;height:50pt' fillcolor='#ff0000'/>";
        _svg = "<svg xmlns='http://www.w3.org/2000/svg'><script>x()</script><rect width='10' height='10' onload='y()'/></svg>";
    }

    [Benchmark]
    public GraphicKind Detect_Emf() => _svc.Detect(_emf, "image/x-emf");

    [Benchmark]
    public int Png_PassThrough()
    {
        int n = 0;
        for (int i = 0; i < GraphicCount; i++)
            n += _svc.ConvertForEditor(new GraphicSource { Data = _png }).Web!.WidthPx;
        return n;
    }

    [Benchmark]
    public int Emf_BlankFallback()
    {
        int n = 0;
        for (int i = 0; i < GraphicCount; i++)
            n += _svc.ConvertForEditor(new GraphicSource { Data = _emf, ContentType = "image/x-emf" }).Web!.HeightPx;
        return n;
    }

    [Benchmark]
    public int Emf_EmbeddedRasterExtraction()
    {
        int n = 0;
        for (int i = 0; i < GraphicCount; i++)
            n += _svc.ConvertForEditor(new GraphicSource { Data = _emfWithPng, ContentType = "image/x-emf" }).Web!.WidthPx;
        return n;
    }

    [Benchmark]
    public int Vml_ToSvg()
    {
        int n = 0;
        for (int i = 0; i < GraphicCount; i++)
            n += _svc.ConvertVmlShapeForEditor(_vml)!.Web!.WidthPx;
        return n;
    }

    [Benchmark]
    public int Svg_Sanitize()
    {
        int n = 0;
        for (int i = 0; i < GraphicCount; i++)
            n += _svc.SanitizeSvg(_svg)!.Length;
        return n;
    }

    // ---- synthetic builders -------------------------------------------------

    private static byte[] MinimalPng(int w, int h)
    {
        var bytes = new List<byte> { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        var ihdr = new byte[25];
        BinaryPrimitives.WriteUInt32BigEndian(ihdr, 13);
        Encoding.ASCII.GetBytes("IHDR").CopyTo(ihdr, 4);
        BinaryPrimitives.WriteUInt32BigEndian(ihdr.AsSpan(8), (uint)w);
        BinaryPrimitives.WriteUInt32BigEndian(ihdr.AsSpan(12), (uint)h);
        ihdr[16] = 8; ihdr[17] = 2;
        bytes.AddRange(ihdr);
        var iend = new byte[12];
        Encoding.ASCII.GetBytes("IEND").CopyTo(iend, 4);
        bytes.AddRange(iend);
        return bytes.ToArray();
    }

    private static byte[] BuildEmf(int frameRight, int frameBottom)
    {
        var d = new byte[88];
        BinaryPrimitives.WriteUInt32LittleEndian(d, 1);
        BinaryPrimitives.WriteUInt32LittleEndian(d.AsSpan(4), 88);
        BinaryPrimitives.WriteInt32LittleEndian(d.AsSpan(32), frameRight);
        BinaryPrimitives.WriteInt32LittleEndian(d.AsSpan(36), frameBottom);
        BinaryPrimitives.WriteUInt32LittleEndian(d.AsSpan(40), 0x464D4520);
        return d;
    }
}
