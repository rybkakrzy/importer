using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Models;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Queries.OpenDocument;

/// <summary>
/// Zapytanie o otwarcie dokumentu DOCX (konwersja na HTML).
/// <paramref name="Password"/> — opcjonalne hasło do odszyfrowania DOCX zabezpieczonego hasłem.
/// </summary>
public record OpenDocumentQuery(Stream FileStream, string FileName, string? Password = null)
    : IRequest<Result<DocumentContent>>;
