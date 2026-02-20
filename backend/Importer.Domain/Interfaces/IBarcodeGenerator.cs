namespace Importer.Domain.Interfaces;

/// <summary>
/// Kontrakt generatora kodów kreskowych i QR
/// </summary>
public interface IBarcodeGenerator
{
    /// <summary>
    /// Generuje kod kreskowy/QR jako obraz PNG
    /// </summary>
    byte[] Generate(string content, string barcodeType, int width, int height, bool showText);

    /// <summary>
    /// Zwraca listę obsługiwanych typów kodów
    /// </summary>
    IReadOnlyList<string> GetSupportedTypes();
}
