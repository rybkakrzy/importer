using D2ViewerEditor.Domain.Models;

namespace D2ViewerEditor.Infrastructure.Services.StructureInspection;

/// <summary>
/// Relationshipy pakietu zaindeksowane parą (część źródłowa, identyfikator). <c>rId1</c> w
/// document.xml i <c>rId1</c> w header1.xml wskazują różne zasoby, więc sam identyfikator nie
/// jest kluczem.
/// </summary>
public sealed class OoxmlRelationshipIndex
{
    private readonly Dictionary<string, StructureRelationship> _relationships;

    public OoxmlRelationshipIndex(Dictionary<string, StructureRelationship> relationships)
    {
        _relationships = relationships;
    }

    public IEnumerable<StructureRelationship> All => _relationships.Values;

    public static string CreateKey(string sourcePart, string relationshipId) => $"{sourcePart}|{relationshipId}";

    public StructureRelationship? Find(string sourcePart, string relationshipId) =>
        _relationships.GetValueOrDefault(CreateKey(sourcePart, relationshipId));

    /// <summary>Pierwszy relationship danego typu wychodzący z części (np. <c>styles</c>, <c>numbering</c>).</summary>
    public string? FindTargetByType(string sourcePart, string typeSuffix) =>
        _relationships.Values
            .Where(relationship =>
                relationship.SourcePart.Equals(sourcePart, StringComparison.OrdinalIgnoreCase) &&
                OoxmlNamespaces.IsRelationshipType(relationship.Type, typeSuffix) &&
                relationship.ResolvedTarget is not null)
            .Select(relationship => relationship.ResolvedTarget)
            .FirstOrDefault();

    /// <summary>Część, z której wychodzi relationship danego typu wskazujący na podaną część.</summary>
    public string? FindSourceByTarget(string targetPath, string typeSuffix) =>
        _relationships.Values
            .Where(relationship =>
                OoxmlNamespaces.IsRelationshipType(relationship.Type, typeSuffix) &&
                relationship.ResolvedTarget?.Equals(targetPath, StringComparison.OrdinalIgnoreCase) == true)
            .Select(relationship => relationship.SourcePart)
            .FirstOrDefault();

    internal void MarkTargetMissing(StructureRelationship relationship)
    {
        _relationships[CreateKey(relationship.SourcePart, relationship.Id)] =
            relationship with { Status = StructureRelationshipStatus.TargetMissing };
    }
}
