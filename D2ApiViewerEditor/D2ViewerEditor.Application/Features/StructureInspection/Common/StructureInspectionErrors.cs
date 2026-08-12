namespace D2ViewerEditor.Application.Features.StructureInspection.Common;

/// <summary>
/// Komunikaty błędów narzędzia. Analiza żyje w pamięci procesu i wygasa — komunikat musi mówić
/// wprost, że trzeba wczytać dokument ponownie, a nie sugerować awarii.
/// </summary>
public static class StructureInspectionErrors
{
    public const string InspectionNotFound =
        "Analiza dokumentu wygasła lub nie istnieje. Wczytaj dokument ponownie.";

    public const string ElementNotFound =
        "Nie znaleziono elementu o podanym identyfikatorze w tej analizie.";

    public const string PartNotFound =
        "Nie znaleziono części XML o podanej ścieżce w pakiecie.";
}
