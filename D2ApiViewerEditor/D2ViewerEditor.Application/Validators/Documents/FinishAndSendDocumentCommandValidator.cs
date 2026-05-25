using D2ViewerEditor.Application.Features.Documents.Commands.FinishAndSendDocument;
using FluentValidation;

namespace D2ViewerEditor.Application.Validators.Documents;

/// <summary>
/// Validator dla FinishAndSendDocumentCommand.
/// </summary>
public class FinishAndSendDocumentCommandValidator : AbstractValidator<FinishAndSendDocumentCommand>
{
    public FinishAndSendDocumentCommandValidator()
    {
        RuleFor(x => x.MasterId)
            .NotEmpty().WithMessage("GUID mastera jest wymagany");

        RuleFor(x => x.VersionId)
            .NotEmpty().WithMessage("GUID wersji jest wymagany");

        RuleFor(x => x.Content)
            .NotNull().WithMessage("Zawartość dokumentu jest wymagana")
            .Must(x => x != null && x.Length > 0).WithMessage("Zawartość dokumentu nie może być pusta")
            .Must(x => x == null || x.Length <= 100 * 1024 * 1024).WithMessage("Rozmiar dokumentu nie może przekraczać 100 MB");

        RuleFor(x => x.CreatedBy)
            .MaximumLength(255).WithMessage("Pole 'CreatedBy' nie może przekraczać 255 znaków");
    }
}
