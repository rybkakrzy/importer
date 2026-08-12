using D2ViewerEditor.Application.Features.StructureInspection.Common;
using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.StructureInspection.Queries.GetStructureSchemaIssues;

/// <summary>
/// Wyniki walidacji schematu. Bez <c>TargetVersion</c> zwracany jest wynik z analizy; wskazanie
/// profilu uruchamia walidację ponownie dla tej wersji Office.
/// </summary>
public record GetStructureSchemaIssuesQuery(Guid InspectionId, string? TargetVersion)
    : IRequest<Result<SchemaIssuesDto>>;
