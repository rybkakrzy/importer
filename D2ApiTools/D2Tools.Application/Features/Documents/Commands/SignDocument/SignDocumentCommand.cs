using D2Tools.Domain.Common;
using D2Tools.Domain.Models;
using MediatR;

namespace D2Tools.Application.Features.Documents.Commands.SignDocument;

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
    string? SignatureReason
) : IRequest<Result<SignDocumentResult>>;

public record SignDocumentResult(byte[] DocxBytes, string FileName);
