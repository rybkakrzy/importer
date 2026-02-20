using Importer.Domain.Common;
using Importer.Domain.Models;
using MediatR;

namespace Importer.Application.Features.Documents.Queries.GetNewDocument;

/// <summary>
/// Zapytanie o nowy pusty dokument
/// </summary>
public record GetNewDocumentQuery : IRequest<Result<DocumentContent>>;
