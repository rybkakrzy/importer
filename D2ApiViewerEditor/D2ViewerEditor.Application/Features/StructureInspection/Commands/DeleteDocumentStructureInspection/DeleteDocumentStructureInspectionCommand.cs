using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.StructureInspection.Commands.DeleteDocumentStructureInspection;

/// <summary>Zwolnienie analizy z magazynu — GUI woła to przy wczytaniu kolejnego dokumentu.</summary>
public record DeleteDocumentStructureInspectionCommand(Guid InspectionId) : IRequest<Result>;
