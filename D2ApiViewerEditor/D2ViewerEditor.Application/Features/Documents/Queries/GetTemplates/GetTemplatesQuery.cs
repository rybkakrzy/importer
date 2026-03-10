using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetTemplates;

/// <summary>
/// Zapytanie o listę szablonów dokumentów
/// </summary>
public record GetTemplatesQuery : IRequest<Result<List<TemplateDto>>>;

public record TemplateDto(string Id, string Name, string Description);
