using D2ViewerEditor.Application.Features.StructureInspection.Common;
using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.StructureInspection.Queries.GetStructureElements;

/// <summary>
/// Spłaszczone drzewo elementów analizy. Filtry są opcjonalne — GUI pobiera listę raz i filtruje
/// lokalnie, ale te same filtry działają po stronie serwera (np. dla wywołań diagnostycznych).
/// </summary>
public record GetStructureElementsQuery(
    Guid InspectionId,
    string? PartPath = null,
    string? Category = null,
    string? Severity = null,
    string? Search = null) : IRequest<Result<List<StructureElementDto>>>;
