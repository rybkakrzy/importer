namespace D2Tools.Domain.Models;

/// <summary>
/// Request do generowania kodu kreskowego / QR
/// </summary>
public class BarcodeRequest
{
    public string Content { get; set; } = string.Empty;
    public string BarcodeType { get; set; } = "QRCode";
    public int Width { get; set; } = 300;
    public int Height { get; set; } = 300;
    public bool ShowText { get; set; } = false;
}

/// <summary>
/// Odpowiedź z wygenerowanym kodem kreskowym
/// </summary>
public class BarcodeResponse
{
    public string Base64Image { get; set; } = string.Empty;
    public string ContentType { get; set; } = "image/png";
    public string BarcodeType { get; set; } = string.Empty;
}
