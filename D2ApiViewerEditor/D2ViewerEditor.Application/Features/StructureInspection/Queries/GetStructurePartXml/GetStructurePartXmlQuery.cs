using D2ViewerEditor.Application.Features.StructureInspection.Common;
using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.StructureInspection.Queries.GetStructurePartXml;

/// <summary>
/// Surowy XML całej części pakietu. <c>HighlightElementId</c> pozwala przejść z konkretnego
/// elementu do jego miejsca w pełnym XML — backend zwraca wtedy numer linii do podświetlenia.
/// </summary>
public record GetStructurePartXmlQuery(Guid InspectionId, string PartPath, string? HighlightElementId)
    : IRequest<Result<StructurePartXmlDto>>;
