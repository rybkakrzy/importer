using BenchmarkDotNet.Running;
using D2ViewerEditor.Benchmarks;

// Two modes:
//   dotnet run -c Release -- --fidelity [--fail-on-regression]   → fast DOCX fidelity report (R-15/16/17/18)
//   dotnet run -c Release -- [BenchmarkDotNet filters]           → performance + memory benchmarks
// Performance benchmarks REQUIRE Release. The fidelity report is quick and Release-agnostic.
if (args.Contains("--fidelity"))
    return FidelityReportRunner.Run(failOnRegression: args.Contains("--fail-on-regression"));

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
return 0;

internal sealed partial class Program;
