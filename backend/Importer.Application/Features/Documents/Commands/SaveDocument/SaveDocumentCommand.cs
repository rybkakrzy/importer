using Importer.Domain.Common;
using Importer.Domain.Models;
using MediatR;

namespace Importer.Application.Features.Documents.Commands.SaveDocument;

/// <summary>
/// Komenda zapisania dokumentu HTML jako DOCX
/// </summary>
public record SaveDocumentCommand(
    string Html,
    string? OriginalFileName,
    DocumentMetadata? Metadata,
    HeaderFooterContent? Header,
    HeaderFooterContent? Footer
) : IRequest<Result<SaveDocumentResult>>;

public record SaveDocumentResult(byte[] DocxBytes, string FileName);
