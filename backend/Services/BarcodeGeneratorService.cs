using SkiaSharp;
using ZXing;
using ZXing.Common;

namespace ImporterApi.Services;

/// <summary>
/// Serwis generowania kodów kreskowych i QR
/// Używa ZXing.Net (cross-platform, działa na Linux/GCP) + SkiaSharp do renderowania
/// </summary>
public class BarcodeGeneratorService
{
    /// <summary>
    /// Obsługiwane typy kodów kreskowych
    /// </summary>
    private static readonly Dictionary<string, BarcodeFormat> SupportedFormats = new(StringComparer.OrdinalIgnoreCase)
    {
        ["QRCode"] = BarcodeFormat.QR_CODE,
        ["Code128"] = BarcodeFormat.CODE_128,
        ["EAN13"] = BarcodeFormat.EAN_13,
        ["UPCA"] = BarcodeFormat.UPC_A,
        ["EAN8"] = BarcodeFormat.EAN_8,
        ["Interleaved2of5"] = BarcodeFormat.ITF,
        ["Code39"] = BarcodeFormat.CODE_39,
        ["AztecCode"] = BarcodeFormat.AZTEC,
        ["Datamatrix"] = BarcodeFormat.DATA_MATRIX,
        ["PDF417"] = BarcodeFormat.PDF_417,
        ["MicroPDF"] = BarcodeFormat.PDF_417, // Micro PDF jako wariant PDF417
        ["Codabar"] = BarcodeFormat.CODABAR,
        ["Code93"] = BarcodeFormat.CODE_93,
        ["Maxicode"] = BarcodeFormat.MAXICODE
    };

    /// <summary>
    /// Pobiera listę obsługiwanych typów kodów
    /// </summary>
    public static IReadOnlyList<string> GetSupportedTypes()
    {
        return SupportedFormats.Keys.ToList().AsReadOnly();
    }

    /// <summary>
    /// Generuje kod kreskowy/QR jako obraz PNG w base64
    /// </summary>
    /// <param name="content">Treść do zakodowania</param>
    /// <param name="barcodeType">Typ kodu</param>
    /// <param name="width">Szerokość w pikselach</param>
    /// <param name="height">Wysokość w pikselach</param>
    /// <returns>Obraz PNG jako tablica bajtów</returns>
    public byte[] GenerateBarcode(string content, string barcodeType, int width = 300, int height = 300)
    {
        if (string.IsNullOrWhiteSpace(content))
            throw new ArgumentException("Treść kodu nie może być pusta", nameof(content));

        if (!SupportedFormats.ContainsKey(barcodeType))
            throw new ArgumentException($"Nieobsługiwany typ kodu: {barcodeType}", nameof(barcodeType));

        var format = SupportedFormats[barcodeType];

        var showHumanReadableText = IsOneDimensional(format);
        var textAreaHeight = showHumanReadableText ? 28 : 0;
        var barcodeHeight = Math.Max(24, height - textAreaHeight);

        // Dla formatów 2D (QR, Aztec, DataMatrix, MaxiCode) - kwadratowe wymiary
        bool is2D = format is BarcodeFormat.QR_CODE or BarcodeFormat.AZTEC
            or BarcodeFormat.DATA_MATRIX or BarcodeFormat.MAXICODE or BarcodeFormat.PDF_417;

        if (is2D && format != BarcodeFormat.PDF_417)
        {
            // Dla kwadratowych 2D - użyj mniejszego wymiaru
            var size = Math.Min(width, height);
            width = size;
            height = size;
        }

        // Konfiguracja encodera
        var writer = new ZXing.BarcodeWriterGeneric
        {
            Format = format,
            Options = new EncodingOptions
            {
                Width = width,
                Height = barcodeHeight,
                Margin = 2,
                PureBarcode = false
            }
        };

        // Generuj matrycę bitową
        var matrix = writer.Encode(content);

        // Renderuj do obrazu PNG za pomocą SkiaSharp
        return RenderToPng(
            matrix,
            width,
            barcodeHeight,
            textAreaHeight,
            showHumanReadableText ? content : null
        );
    }

    private static bool IsOneDimensional(BarcodeFormat format)
    {
        return format is BarcodeFormat.CODE_128
            or BarcodeFormat.EAN_13
            or BarcodeFormat.UPC_A
            or BarcodeFormat.EAN_8
            or BarcodeFormat.ITF
            or BarcodeFormat.CODE_39
            or BarcodeFormat.CODABAR
            or BarcodeFormat.CODE_93;
    }

    /// <summary>
    /// Renderuje BitMatrix do PNG za pomocą SkiaSharp (cross-platform)
    /// </summary>
    private byte[] RenderToPng(
        BitMatrix matrix,
        int requestedWidth,
        int barcodeHeight,
        int textAreaHeight,
        string? humanReadableText)
    {
        var matrixWidth = matrix.Width;
        var matrixHeight = matrix.Height;
        var totalHeight = barcodeHeight + textAreaHeight;

        // Użyj wymiarów matrycy, ale skaluj do żądanych rozmiarów
        var scaleX = (float)requestedWidth / matrixWidth;
        var scaleY = (float)barcodeHeight / matrixHeight;
        var scale = Math.Min(scaleX, scaleY);

        var drawWidth = matrixWidth * scale;
        var drawHeight = matrixHeight * scale;
        var offsetX = (requestedWidth - drawWidth) / 2f;
        var offsetY = (barcodeHeight - drawHeight) / 2f;

        using var bitmap = new SKBitmap(requestedWidth, totalHeight);
        using var canvas = new SKCanvas(bitmap);

        // Białe tło
        canvas.Clear(SKColors.White);

        using var blackPaint = new SKPaint
        {
            Color = SKColors.Black,
            Style = SKPaintStyle.Fill,
            IsAntialias = false
        };

        // Rysuj piksele kodu
        for (int y = 0; y < matrixHeight; y++)
        {
            for (int x = 0; x < matrixWidth; x++)
            {
                if (matrix[x, y])
                {
                    canvas.DrawRect(
                        offsetX + x * scale,
                        offsetY + y * scale,
                        scale,
                        scale,
                        blackPaint
                    );
                }
            }
        }

        if (!string.IsNullOrWhiteSpace(humanReadableText) && textAreaHeight > 0)
        {
            using var textPaint = new SKPaint
            {
                Color = SKColors.Black,
                IsAntialias = true,
                TextAlign = SKTextAlign.Center,
                Typeface = SKTypeface.FromFamilyName("Arial")
            };

            textPaint.TextSize = 14;
            while (textPaint.MeasureText(humanReadableText) > requestedWidth - 8 && textPaint.TextSize > 8)
            {
                textPaint.TextSize -= 1;
            }

            var baselineY = barcodeHeight + textAreaHeight - 8;
            canvas.DrawText(humanReadableText, requestedWidth / 2f, baselineY, textPaint);
        }

        // Koduj do PNG
        using var image = SKImage.FromBitmap(bitmap);
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);

        return data.ToArray();
    }
}
