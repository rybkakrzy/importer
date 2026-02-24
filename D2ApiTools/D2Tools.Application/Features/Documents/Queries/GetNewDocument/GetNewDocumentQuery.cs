using D2Tools.Domain.Common;
using D2Tools.Domain.Models;
using MediatR;

namespace D2Tools.Application.Features.Documents.Queries.GetNewDocument;

/// <summary>
/// Zapytanie o nowy pusty dokument
/// </summary>
public record GetNewDocumentQuery : IRequest<Result<DocumentContent>>;
