using D2ViewerEditor.Application.Features.Documents.Commands.UploadDocument;
using FluentValidation;

namespace D2ViewerEditor.Application.Validators.Documents;

/// <summary>
/// Validator dla UploadDocumentCommand
/// </summary>
public class UploadDocumentCommandValidator : AbstractValidator<UploadDocumentCommand>
{
    public UploadDocumentCommandValidator()
    {
        RuleFor(x => x.FileName)
            .NotEmpty().WithMessage("Nazwa dokumentu jest wymagana")
            .MaximumLength(500).WithMessage("Nazwa dokumentu nie może przekraczać 500 znaków");

        RuleFor(x => x.MimeType)
            .NotEmpty().WithMessage("Typ MIME jest wymagany")
            .MaximumLength(255).WithMessage("Typ MIME nie może przekraczać 255 znaków")
            .Must(BeValidMimeType).WithMessage("Nieprawidłowy format typu MIME");

        RuleFor(x => x.Content)
            .NotNull().WithMessage("Zawartość dokumentu jest wymagana")
            .Must(x => x != null && x.Length > 0).WithMessage("Zawartość dokumentu nie może być pusta")
            .Must(x => x == null || x.Length <= 100 * 1024 * 1024).WithMessage("Rozmiar dokumentu nie może przekraczać 100 MB");

        RuleFor(x => x.CreatedBy)
            .NotEmpty().WithMessage("Pole 'CreatedBy' jest wymagane")
            .MaximumLength(255).WithMessage("Pole 'CreatedBy' nie może przekraczać 255 znaków");
    }

    private bool BeValidMimeType(string mimeType)
    {
        if (string.IsNullOrWhiteSpace(mimeType))
            return false;

        // Basic MIME type validation: type/subtype
        var parts = mimeType.Split('/');
        return parts.Length == 2 && 
               !string.IsNullOrWhiteSpace(parts[0]) && 
               !string.IsNullOrWhiteSpace(parts[1]);
    }
}
