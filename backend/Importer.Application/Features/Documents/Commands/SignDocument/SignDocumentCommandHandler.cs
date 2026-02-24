using Importer.Domain.Common;
using Importer.Domain.Interfaces;
using MediatR;

namespace Importer.Application.Features.Documents.Commands.SignDocument;

public class SignDocumentCommandHandler : IRequestHandler<SignDocumentCommand, Result<SignDocumentResult>>
{
    private readonly IHtmlToDocxConverter _converter;
    private readonly IDigitalSignatureService _signatureService;

    public SignDocumentCommandHandler(IHtmlToDocxConverter converter, IDigitalSignatureService signatureService)
    {
        _converter = converter;
        _signatureService = signatureService;
    }

    public Task<Result<SignDocumentResult>> Handle(SignDocumentCommand request, CancellationToken cancellationToken)
    {
        try
        {
            // Konwertuj HTML na DOCX
            var docxBytes = _converter.Convert(request.Html, request.Metadata, request.Header, request.Footer);

            // Zdekoduj certyfikat
            var certificateBytes = Convert.FromBase64String(request.CertificateBase64);

            // Podpisz dokument
            var signedBytes = _signatureService.SignDocument(
                docxBytes,
                certificateBytes,
                request.CertificatePassword,
                request.SignerName,
                request.SignerTitle,
                request.SignerEmail,
                request.SignatureReason
            );

            var fileName = !string.IsNullOrEmpty(request.OriginalFileName)
                ? request.OriginalFileName
                : "dokument.docx";

            if (!fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                fileName += ".docx";

            return Task.FromResult(Result<SignDocumentResult>.Success(new SignDocumentResult(signedBytes, fileName)));
        }
        catch (Exception ex)
        {
            return Task.FromResult(Result<SignDocumentResult>.Failure($"Błąd podpisywania dokumentu: {ex.Message}"));
        }
    }
}
