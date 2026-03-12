using D2ViewerEditor.Application.Features.Documents.Commands.SaveDocumentVersion;
using FluentValidation;

namespace D2ViewerEditor.Application.Validators.Documents;

/// <summary>
/// Validator dla SaveDocumentVersionCommand
/// </summary>
public class SaveDocumentVersionCommandValidator : AbstractValidator<SaveDocumentVersionCommand>
{
    public SaveDocumentVersionCommandValidator()
    {
        RuleFor(x => x.MasterId)
            .NotEmpty().WithMessage("GUID mastera jest wymagany");

        RuleFor(x => x.Content)
            .NotNull().WithMessage("Zawartość dokumentu jest wymagana")
            .Must(x => x != null && x.Length > 0).WithMessage("Zawartość dokumentu nie może być pusta")
            .Must(x => x == null || x.Length <= 100 * 1024 * 1024).WithMessage("Rozmiar dokumentu nie może przekraczać 100 MB");

        RuleFor(x => x.CreatedBy)
            .NotEmpty().WithMessage("Pole 'CreatedBy' jest wymagane")
            .MaximumLength(255).WithMessage("Pole 'CreatedBy' nie może przekraczać 255 znaków");
    }
}
