using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Models;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.SaveDocument;

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
