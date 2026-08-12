namespace D2ViewerEditor.Infrastructure.Services.StructureInspection;

/// <summary>
/// Profil obsługi cech OOXML przez nasz edytor. To DANE konfiguracji — walidator nie wnioskuje
/// poziomów z kodu edytora. Cecha nieopisana w profilu dostaje <see cref="DefaultLevel"/>.
/// Dopuszczalne poziomy: Supported, Partial, Unsupported, Unknown.
/// </summary>
public sealed class EditorCompatibilityOptions
{
    public const string SectionName = "EditorCompatibility";

    public string ProfileName { get; set; } = "Nieskonfigurowany";
    public string DefaultLevel { get; set; } = "Unknown";
    public Dictionary<string, string> Features { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}
