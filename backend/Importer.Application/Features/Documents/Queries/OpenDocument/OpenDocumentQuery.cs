using Importer.Domain.Common;
using Importer.Domain.Models;
using MediatR;

namespace Importer.Application.Features.Documents.Queries.OpenDocument;

/// <summary>
/// Zapytanie o otwarcie dokumentu DOCX (konwersja na HTML)
/// </summary>
public record OpenDocumentQuery(Stream FileStream, string FileName) : IRequest<Result<DocumentContent>>;
