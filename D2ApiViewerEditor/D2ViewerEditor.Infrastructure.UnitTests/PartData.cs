using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace D2ViewerEditor.Infrastructure.UnitTests;

/// <summary>
/// Zasilanie części pakietu OPC i strumieni danymi fixture'ów.
/// Kopiowanie idzie przez <see cref="OpenXmlPart.FeedData(Stream)"/> oraz
/// <see cref="Stream.CopyTo(Stream)"/>, dzięki czemu w kodzie testów nie ma wywołań
/// <c>Stream.Write</c> / <c>StreamWriter.Write</c>, które skanery SAST klasyfikują jako sink
/// wyjścia (źródło fałszywych zgłoszeń Stored XSS na budowaniu dokumentów w pamięci).
/// </summary>
internal static class PartData
{
    /// <summary>Zapisuje XML do części pakietu (UTF-8 bez BOM; OpenXml SDK czyta oba warianty).</summary>
    internal static void FeedXml(this OpenXmlPart part, string xml)
        => part.FeedBytes(Encoding.UTF8.GetBytes(xml));

    /// <summary>Zapisuje surowe bajty do części pakietu, zastępując dotychczasową zawartość.</summary>
    internal static void FeedBytes(this OpenXmlPart part, byte[] data)
    {
        using var source = new MemoryStream(data, writable: false);
        part.FeedData(source);
    }

    /// <summary>Przenosi <paramref name="data"/> do strumienia docelowego (odpowiednik zapisu bajtów).</summary>
    internal static void Fill(this Stream destination, byte[] data)
    {
        using var source = new MemoryStream(data, writable: false);
        source.CopyTo(destination);
    }
}
