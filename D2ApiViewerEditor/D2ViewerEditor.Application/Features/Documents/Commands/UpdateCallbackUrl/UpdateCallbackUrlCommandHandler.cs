using D2ViewerEditor.Application.Features.Documents.Common;
using D2ViewerEditor.Application.Common.Security;
using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Entities;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.UpdateCallbackUrl;

public class UpdateCallbackUrlCommandHandler
    : IRequestHandler<UpdateCallbackUrlCommand, Result>
{
    private readonly IDocumentRepository _documentRepository;
    private readonly IReturnUrlValidator _returnUrlValidator;

    public UpdateCallbackUrlCommandHandler(IDocumentRepository documentRepository, IReturnUrlValidator returnUrlValidator)
    {
        _documentRepository = documentRepository;
        _returnUrlValidator = returnUrlValidator;
    }

    public async Task<Result> Handle(UpdateCallbackUrlCommand request, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.CallbackUrl))
            return Result.Failure("Callback URL jest wymagany.");

        var urlValidation = _returnUrlValidator.Validate(request.CallbackUrl);
        if (!urlValidation.IsValid)
            return Result.Failure(urlValidation.Error!);

        var document = await _documentRepository.GetByIdAsync(request.MasterId, cancellationToken);
        if (document == null)
            return Result.NotFound();

        if (document.Status is DocumentStatus.Sending or DocumentStatus.Sent or DocumentStatus.DeliveryFailed)
            return Result.Failure(
                $"Nie można zaktualizować callback URL dla dokumentu w stanie {document.Status}: wysyłka już została zlecona.");

        // Preserve every other field on the metadata blob — we only own ReturnUrl here.
        var existing = ExternalDocumentMetadata.Parse(document.Metadata);
        var newMetadata = (existing with { ReturnUrl = urlValidation.NormalizedUrl }).Serialize();

        document.UpdateMetadata(newMetadata);
        await _documentRepository.SaveChangesAsync(cancellationToken);

        return Result.Success();
    }
}
