using D2Tools.Domain.Common;
using D2Tools.Domain.Models;
using MediatR;

namespace D2Tools.Application.Features.Documents.Queries.OpenDocument;

/// <summary>
/// Zapytanie o otwarcie dokumentu DOCX (konwersja na HTML)
/// </summary>
public record OpenDocumentQuery(Stream FileStream, string FileName) : IRequest<Result<DocumentContent>>;
