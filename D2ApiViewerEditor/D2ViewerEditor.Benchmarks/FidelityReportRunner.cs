using System.Text;

namespace D2ViewerEditor.Benchmarks;

/// <summary>
/// Runs the DOCX → DocumentContent → DOCX round-trip on the benchmark samples and reports
/// whether the document's fidelity invariants held (R-15/R-16/R-17/R-18). The original document
/// is its own baseline — every metric is original-vs-result, so no stored baseline is needed.
/// Report-only by default; pass <c>--fail-on-regression</c> to return a non-zero exit on any FAIL.
/// </summary>
public static class FidelityReportRunner
{
    public static int Run(bool failOnRegression, string? outputDirectory = null)
    {
        var samples = new List<(string Name, byte[] Original, bool PassThrough)>
        {
            ("simple", BenchmarkAssets.Simple(), false),
            ("tables", BenchmarkAssets.Tables(), true),
            ("images", BenchmarkAssets.Images(), false),
            ("header-footer", BenchmarkAssets.HeaderFooter(), false),
            ("page-breaks", BenchmarkAssets.PageBreaks(), false),
            ("multi-style (pass-through)", BenchmarkAssets.MultiStyle(), true),
        };

        var regression = BenchmarkAssets.TryLoadRegressionOriginal();
        if (regression != null)
            samples.Add(($"regression: {BenchmarkAssets.RegressionOriginalFileName} (pass-through)", regression, true));

        var sections = new StringBuilder();
        var worst = FidelityStatus.Pass;

        foreach (var (name, original, passThrough) in samples)
        {
            var content = Pipeline.Read(original);
            var result = passThrough ? Pipeline.WritePreserving(content, original) : Pipeline.Write(content);

            var before = DocxPackageReport.Analyze(original);
            var after = DocxPackageReport.Analyze(result, content.Html.Length);
            var cmp = FidelityComparison.Compare(before, after, preservationExpected: passThrough);
            if (cmp.Overall > worst) worst = cmp.Overall;

            sections.AppendLine($"## {name} — **{cmp.Overall.ToString().ToUpperInvariant()}**");
            sections.AppendLine();
            sections.AppendLine($"HTML length: {content.Html.Length} · result size: {after.FileSizeBytes} B · parts {before.PartCount}→{after.PartCount}");
            sections.AppendLine();
            sections.AppendLine("| Check | Status | Detail |");
            sections.AppendLine("|---|---|---|");
            foreach (var c in cmp.Checks)
                sections.AppendLine($"| {c.Name} | {Badge(c.Status)} | {c.Detail} |");
            sections.AppendLine();
        }

        var sb = new StringBuilder();
        sb.AppendLine("# DOCX fidelity report");
        sb.AppendLine();
        sb.AppendLine($"**Overall: {worst.ToString().ToUpperInvariant()}**");
        sb.AppendLine();
        sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}  ·  assets dir: `{BenchmarkAssets.AssetsDirectory}`");
        sb.AppendLine();
        sb.Append(sections);

        var report = sb.ToString();
        Console.WriteLine(report);

        var dir = outputDirectory ?? Path.Combine(AppContext.BaseDirectory, "fidelity-reports");
        Directory.CreateDirectory(dir);
        var file = Path.Combine(dir, $"fidelity-{DateTime.Now:yyyyMMdd-HHmmss}.md");
        File.WriteAllText(file, report);
        Console.WriteLine($"Report written: {file}");

        if (failOnRegression && worst == FidelityStatus.Fail)
        {
            Console.Error.WriteLine("Fidelity FAIL detected (--fail-on-regression).");
            return 1;
        }
        return 0;
    }

    private static string Badge(FidelityStatus s) => s switch
    {
        FidelityStatus.Pass => "PASS",
        FidelityStatus.Warn => "WARN",
        _ => "FAIL"
    };
}
