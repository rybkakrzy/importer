using D2Tools.Domain.Common;
using D2Tools.Domain.Models;
using MediatR;

namespace D2Tools.Application.Features.Documents.Commands.SaveDocument;

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
