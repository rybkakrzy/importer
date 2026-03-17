using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using D2ViewerEditor.Infrastructure.Services;

namespace D2ViewerEditor.Benchmarks.Benchmarks;

/// <summary>
/// Benchmarki generowania kodów kreskowych i QR.
///
/// Mierzone obszary:
///  - Encoding przez ZXing.Net (budowanie BitMatrix)
///  - Renderowanie piksel po pikselu przez SkiaSharp (pętla DrawRect po BitMatrix)
///  - Kodowanie finalnego PNG
///
/// Kluczowy hot-path: metoda RenderToPng wykonuje pętlę O(W×H) wywołań DrawRect.
/// Przy rozmiarze 600×600 = 360 000 potencjalnych wywołań rysowania.
/// </summary>
[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
public class BarcodeBenchmarks
{
    private BarcodeGeneratorService _service = null!;

    [GlobalSetup]
    public void Setup()
    {
        _service = new BarcodeGeneratorService();
    }

    // --- Parametr rozmiaru (QR jest kwadratowy, więc Width == Height) ---

    [Params(150, 300, 600)]
    public int Size { get; set; }

    // --- QR Code ---

    /// <summary>
    /// Klasyczny QR Code. Format 2D — pętla DrawRect po macierzy N×N.
    /// Rozmiar matrycy zależy od długości danych i poziomu korekcji błędów.
    /// </summary>
    [Benchmark]
    public byte[] QrCode()
        => _service.Generate("https://example.com/benchmark-test-2024", "QRCode", Size, Size);

    /// <summary>
    /// QR Code z dłuższą treścią — wymusza większą matrycę, więcej modułów do narysowania.
    /// </summary>
    [Benchmark]
    public byte[] QrCodeLongContent()
        => _service.Generate(
            "Numer zamówienia: PL/2024/ABC/123456789 | Kontrahent: Firma Przykładowa Sp. z o.o. | Kwota: 12 345,67 PLN",
            "QRCode", Size, Size);

    // --- Kody 1D ---

    /// <summary>
    /// Code128 bez tekstu pod spodem — renderuje tylko piksele kodu.
    /// </summary>
    [Benchmark(Baseline = true)]
    public byte[] Code128()
        => _service.Generate("BENCHMARK123456", "Code128", Size, /* height */ 100);

    /// <summary>
    /// Code128 z tekstem — uruchamia pętlę skalowania czcionki + DrawText.
    /// </summary>
    [Benchmark]
    public byte[] Code128WithText()
        => _service.Generate("BENCHMARK123456", "Code128", Size, 100, showText: true);

    /// <summary>
    /// EAN-13 — stały format 13 cyfr, znormalizowany rozmiar matrycy.
    /// </summary>
    [Benchmark]
    public byte[] Ean13()
        => _service.Generate("5901234123457", "EAN13", Size, 100);

    // --- Kody 2D (inne niż QR) ---

    /// <summary>
    /// PDF417 — format 2D, ale prostokątny (nie kwadratowy).
    /// Generuje większe matryce niż QR przy tej samej treści.
    /// </summary>
    [Benchmark]
    public byte[] Pdf417()
        => _service.Generate("Benchmark PDF417 content — dane dokumentu transportowego 2024", "PDF417", Size, Size);

    /// <summary>
    /// Data Matrix — kompaktowy format 2D, mała matryca dla krótkich treści.
    /// </summary>
    [Benchmark]
    public byte[] DataMatrix()
        => _service.Generate("DM-BenchmarkData", "Datamatrix", Size, Size);
}
