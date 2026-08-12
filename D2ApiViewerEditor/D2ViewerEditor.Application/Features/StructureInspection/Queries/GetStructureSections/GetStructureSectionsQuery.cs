using D2ViewerEditor.Application.Features.StructureInspection.Common;
using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.StructureInspection.Queries.GetStructureSections;

/// <summary>Sekcje dokumentu wraz z efektywnymi wiązaniami nagłówków i stopek.</summary>
public record GetStructureSectionsQuery(Guid InspectionId) : IRequest<Result<List<DocumentSectionDto>>>;
