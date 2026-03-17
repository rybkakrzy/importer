using BenchmarkDotNet.Running;

// Uruchom wszystkie benchmarki lub wybierz konkretne używając interaktywnego menu BenchmarkDotNet.
// Tryb Release jest wymagany: dotnet run -c Release
BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
