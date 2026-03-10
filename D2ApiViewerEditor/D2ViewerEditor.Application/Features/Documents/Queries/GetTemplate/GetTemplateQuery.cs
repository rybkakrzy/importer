using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Models;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetTemplate;

/// <summary>
/// Zapytanie o konkretny szablon dokumentu
/// </summary>
public record GetTemplateQuery(string TemplateId) : IRequest<Result<DocumentContent>>;
