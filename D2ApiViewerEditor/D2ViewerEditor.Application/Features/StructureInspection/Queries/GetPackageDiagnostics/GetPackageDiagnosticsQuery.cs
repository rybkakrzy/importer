using D2ViewerEditor.Application.Features.StructureInspection.Common;
using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.StructureInspection.Queries.GetPackageDiagnostics;

/// <summary>Diagnostyka pakietu OPC oraz lista profili walidacji schematu.</summary>
public record GetPackageDiagnosticsQuery(Guid InspectionId) : IRequest<Result<PackageDiagnosticsDto>>;
