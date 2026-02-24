using D2Tools.Domain.Common;
using D2Tools.Domain.Models;
using MediatR;

namespace D2Tools.Application.Features.Documents.Queries.GetTemplate;

/// <summary>
/// Zapytanie o konkretny szablon dokumentu
/// </summary>
public record GetTemplateQuery(string TemplateId) : IRequest<Result<DocumentContent>>;
