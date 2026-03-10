using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Models;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.GetNewDocument;

/// <summary>
/// Zapytanie o nowy pusty dokument
/// </summary>
public record GetNewDocumentQuery : IRequest<Result<DocumentContent>>;
