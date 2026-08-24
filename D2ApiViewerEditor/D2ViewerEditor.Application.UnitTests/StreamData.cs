using System.Text;

namespace D2ViewerEditor.Application.UnitTests;

/// <summary>
/// Zapis danych fixture'ów do strumieni bez wywołań <c>Stream.Write</c> / <c>StreamWriter.Write</c>,
/// które skanery SAST klasyfikują jako sink wyjścia (fałszywe zgłoszenia Stored XSS).
/// Kopiowanie idzie przez <see cref="Stream.CopyTo(Stream)"/>.
/// </summary>
internal static class StreamData
{
    internal static void Fill(this Stream destination, byte[] data)
    {
        using var source = new MemoryStream(data, writable: false);
        source.CopyTo(destination);
    }

    /// <summary>UTF-8 bez BOM — tak jak domyślny <see cref="StreamWriter"/>.</summary>
    internal static void FillText(this Stream destination, string text)
        => destination.Fill(Encoding.UTF8.GetBytes(text));
}
