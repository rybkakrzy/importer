using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace D2ViewerEditor.Benchmarks;

/// <summary>
/// Zasilanie części pakietu OPC danymi assetów benchmarkowych — odpowiednik pomocnika z projektu testów.
/// Kopiowanie przez <see cref="OpenXmlPart.FeedData(Stream)"/> zamiast <c>Stream.Write</c>.
/// </summary>
internal static class PartData
{
    internal static void FeedXml(this OpenXmlPart part, string xml)
        => part.FeedBytes(Encoding.UTF8.GetBytes(xml));

    internal static void FeedBytes(this OpenXmlPart part, byte[] data)
    {
        using var source = new MemoryStream(data, writable: false);
        part.FeedData(source);
    }
}
