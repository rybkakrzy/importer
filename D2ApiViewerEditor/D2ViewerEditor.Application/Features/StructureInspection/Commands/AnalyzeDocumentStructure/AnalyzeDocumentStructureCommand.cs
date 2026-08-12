using D2ViewerEditor.Application.Features.StructureInspection.Common;
using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.StructureInspection.Commands.AnalyzeDocumentStructure;

/// <summary>
/// Wczytanie DOCX do narzędzia „Walidator struktury" i wykonanie analizy. Wynik jest zapamiętywany
/// pod identyfikatorem, po którym pobiera się szczegóły elementów i surowy XML.
/// </summary>
public record AnalyzeDocumentStructureCommand(Stream FileStream, string FileName)
    : IRequest<Result<StructureInspectionSummaryDto>>;
