namespace ImporterApi.Models;

/// <summary>
/// Request do generowania kodu kreskowego / QR
/// </summary>
public class BarcodeRequest
{
    /// <summary>
    /// Treść do zakodowania
    /// </summary>
    public string Content { get; set; } = string.Empty;

    /// <summary>
    /// Typ kodu (QRCode, Code128, EAN13, itp.)
    /// </summary>
    public string BarcodeType { get; set; } = "QRCode";

    /// <summary>
    /// Szerokość obrazu w pikselach (domyślnie 300)
    /// </summary>
    public int Width { get; set; } = 300;

    /// <summary>
    /// Wysokość obrazu w pikselach (domyślnie 300)
    /// </summary>
    public int Height { get; set; } = 300;
}

/// <summary>
/// Odpowiedź z wygenerowanym kodem kreskowym
/// </summary>
public class BarcodeResponse
{
    /// <summary>
    /// Obraz w formacie base64 (data URI)
    /// </summary>
    public string Base64Image { get; set; } = string.Empty;

    /// <summary>
    /// Typ MIME obrazu
    /// </summary>
    public string ContentType { get; set; } = "image/png";

    /// <summary>
    /// Typ kodu który został wygenerowany
    /// </summary>
    public string BarcodeType { get; set; } = string.Empty;
}
