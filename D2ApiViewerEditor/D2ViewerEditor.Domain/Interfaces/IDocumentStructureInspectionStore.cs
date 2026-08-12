using D2ViewerEditor.Domain.Models;

namespace D2ViewerEditor.Domain.Interfaces;

/// <summary>
/// Krótkotrwałe przechowanie analizy struktury dokumentu między wczytaniem pliku a leniwym
/// pobieraniem XML elementu/części. Magazyn jest ograniczony rozmiarem i czasem życia — to
/// narzędzie diagnostyczne, a nie trwały zasób dokumentowy.
/// </summary>
public interface IDocumentStructureInspectionStore
{
    void Save(DocumentStructureInspection inspection);

    DocumentStructureInspection? Get(Guid inspectionId);

    /// <summary>Usuwa analizę; zwraca <c>false</c>, gdy nic nie było pod tym identyfikatorem.</summary>
    bool Delete(Guid inspectionId);
}
