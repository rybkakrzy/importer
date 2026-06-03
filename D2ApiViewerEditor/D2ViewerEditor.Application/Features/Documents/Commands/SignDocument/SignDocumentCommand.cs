using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Models;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.SignDocument;

/// <summary>
/// Komenda podpisania dokumentu certyfikatem cyfrowym
/// </summary>
public record SignDocumentCommand(
    string Html,
    string? OriginalFileName,
    DocumentMetadata? Metadata,
    HeaderFooterContent? Header,
    HeaderFooterContent? Footer,
    string CertificateBase64,
    string CertificatePassword,
    string SignerName,
    string? SignerTitle,
    string? SignerEmail,
    string? SignatureReason,
    PageMargins? Margins = null,
    PageSize? PageSize = null
) : IRequest<Result<SignDocumentResult>>;

public record SignDocumentResult(byte[] DocxBytes, string FileName);
