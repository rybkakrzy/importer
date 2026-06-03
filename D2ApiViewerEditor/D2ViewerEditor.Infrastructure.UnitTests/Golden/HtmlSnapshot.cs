using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Golden;

/// <summary>
/// Minimal approval-style snapshot helper for converter HTML output.
///
/// The snapshot locks the *current* rendered HTML so later refactors (intermediate
/// model, style resolver, layout) cannot silently change the output. It is a
/// regression guard, not a statement that the current output is pixel-correct.
///
/// Volatile fragments are normalized away: base64 image payloads (kept stable but
/// compacted) and date-field text. Baselines live next to the test source under
/// __snapshots__/ and are created automatically on first run (then committed).
/// </summary>
public static class HtmlSnapshot
{
    public static void Verify(string actual, string snapshotName, [CallerFilePath] string callerPath = "")
    {
        var dir = Path.Combine(Path.GetDirectoryName(callerPath)!, "__snapshots__");
        Directory.CreateDirectory(dir);
        var file = Path.Combine(dir, snapshotName + ".approved.html");
        var normalized = Normalize(actual);

        if (!File.Exists(file))
        {
            File.WriteAllText(file, normalized);
            TestContext.Out.WriteLine($"Snapshot baseline created: {file}. Commit it and re-run to lock.");
            return;
        }

        var expected = File.ReadAllText(file).Replace("\r\n", "\n");
        if (normalized != expected)
        {
            Assert.Fail(
                $"HTML snapshot '{snapshotName}' changed.\n" +
                $"--- expected (approved) ---\n{expected}\n" +
                $"--- actual ---\n{normalized}\n" +
                $"If the change is intended, delete {file} and re-run to regenerate.");
        }
    }

    public static string Normalize(string html)
    {
        if (string.IsNullOrEmpty(html)) return string.Empty;

        // Compact base64 image payloads to a stable, short marker (length is content-stable).
        html = Regex.Replace(html, "(data:[^;]+;base64,)[A-Za-z0-9+/=]+",
            m => $"{m.Groups[1].Value}[base64]");

        // Date fields render with DateTime.Now — replace their text content.
        html = Regex.Replace(html, "(<span class=\"field-date\">)[^<]*", "$1[date]");

        // Image relationship ids (data-image-id) are randomized per OOXML build.
        html = Regex.Replace(html, "(data-image-id=\")[^\"]*", "$1[id]");

        // Readable line-based diffs: one tag boundary per line.
        html = html.Replace("\r\n", "\n").Replace("><", ">\n<");

        return html.TrimEnd();
    }
}
