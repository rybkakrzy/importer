using D2ViewerEditor.Application.Features.StructureInspection.Common;
using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.StructureInspection.Queries.GetStructureElementDetails;

/// <summary>Szczegóły jednego elementu: atrybuty, właściwości semantyczne, relationshipy, diagnostyka.</summary>
public record GetStructureElementDetailsQuery(Guid InspectionId, string ElementId)
    : IRequest<Result<StructureElementDetailsDto>>;
