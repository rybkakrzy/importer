using Importer.Domain.Common;
using Importer.Domain.Models;
using MediatR;

namespace Importer.Application.Features.Documents.Queries.GetTemplate;

/// <summary>
/// Zapytanie o konkretny szablon dokumentu
/// </summary>
public record GetTemplateQuery(string TemplateId) : IRequest<Result<DocumentContent>>;
