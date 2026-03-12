using D2ViewerEditor.Application.Features.Documents.Commands.RestoreDocumentVersion;
using FluentValidation;

namespace D2ViewerEditor.Application.Validators.Documents;

/// <summary>
/// Validator dla RestoreDocumentVersionCommand
/// </summary>
public class RestoreDocumentVersionCommandValidator : AbstractValidator<RestoreDocumentVersionCommand>
{
    public RestoreDocumentVersionCommandValidator()
    {
        RuleFor(x => x.MasterId)
            .NotEmpty().WithMessage("GUID mastera jest wymagany");

        RuleFor(x => x.VersionId)
            .NotEmpty().WithMessage("GUID wersji jest wymagany");
    }
}
