namespace D2ViewerEditor.Application.Features.StructureInspection.Common;

/// <summary>
/// Czas życia analizy w magazynie. Trzymany w warstwie aplikacji, bo to reguła użycia narzędzia,
/// a nie szczegół odczytu pakietu; wartość pochodzi z tej samej sekcji konfiguracji co limity.
/// </summary>
public sealed class StructureInspectionRetentionOptions
{
    public const string SectionName = "StructureInspection";

    public TimeSpan InspectionTimeToLive { get; set; } = TimeSpan.FromMinutes(30);
}
