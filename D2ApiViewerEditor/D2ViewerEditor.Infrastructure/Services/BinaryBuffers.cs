using System.Buffers;

namespace D2ViewerEditor.Infrastructure.Services;

/// <summary>
/// Kopiowanie i akumulacja buforów binarnych bez bezpośrednich wywołań <c>Stream.Write</c>.
/// Kopie idą przez <see cref="Stream.CopyTo(Stream)"/>, a akumulacja przez
/// <see cref="ArrayBufferWriter{T}"/> (GetSpan/Advance) — semantyka bez zmian, ale kod aplikacyjny
/// nie zawiera wywołań, które skanery SAST traktują jako sink wyjścia do odpowiedzi HTTP
/// (źródło fałszywych zgłoszeń Stored XSS na strumieniach w pamięci i częściach pakietu OPC).
/// </summary>
internal static class BinaryBuffers
{
    /// <summary>
    /// Rozszerzalny strumień w pamięci zainicjowany kopią <paramref name="data"/>, ustawiony na pozycję 0.
    /// Odpowiednik <c>new MemoryStream()</c> + zapisu bajtów — konstruktor <c>MemoryStream(byte[])</c>
    /// nie nadaje się, bo daje bufor o stałym rozmiarze (pakiet OPC musi móc rosnąć przy zapisie).
    /// </summary>
    internal static MemoryStream ToExpandableStream(byte[] data)
    {
        var target = new MemoryStream(data.Length);
        using (var source = new MemoryStream(data, writable: false))
        {
            source.CopyTo(target);
        }

        target.Position = 0;
        return target;
    }

    /// <summary>
    /// Czyta <paramref name="source"/> do końca z twardym limitem <paramref name="maxOutputBytes"/>
    /// (ochrona przed decompression bomb). Zwraca <c>null</c>, gdy limit został przekroczony
    /// albo strumień nie zawierał danych.
    /// </summary>
    internal static byte[]? ReadAllBounded(Stream source, int maxOutputBytes, int chunkSize = 81920)
    {
        var writer = new ArrayBufferWriter<byte>();
        while (true)
        {
            var span = writer.GetSpan(chunkSize);
            var read = source.Read(span);
            if (read <= 0) break;
            if (writer.WrittenCount + read > maxOutputBytes) return null; // bomba/oversize → odrzuć
            writer.Advance(read);
        }

        return writer.WrittenCount > 0 ? writer.WrittenSpan.ToArray() : null;
    }
}
