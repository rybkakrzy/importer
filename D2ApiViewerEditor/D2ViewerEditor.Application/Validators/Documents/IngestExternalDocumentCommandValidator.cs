using D2ViewerEditor.Application.Features.Documents.Commands.IngestExternalDocument;
using FluentValidation;

namespace D2ViewerEditor.Application.Validators.Documents;

public class IngestExternalDocumentCommandValidator : AbstractValidator<IngestExternalDocumentCommand>
{
    public IngestExternalDocumentCommandValidator()
    {
        RuleFor(x => x.FileName)
            .NotEmpty().WithMessage("Nazwa dokumentu jest wymagana")
            .MaximumLength(500).WithMessage("Nazwa dokumentu nie może przekraczać 500 znaków");

        RuleFor(x => x.MimeType)
            .NotEmpty().WithMessage("Typ MIME jest wymagany")
            .Must(BeSupportedMimeType)
            .WithMessage("Wspierane są tylko pliki DOCX i PDF");

        RuleFor(x => x.Content)
            .NotNull().WithMessage("Zawartość dokumentu jest wymagana")
            .Must(x => x != null && x.Length > 0).WithMessage("Zawartość dokumentu nie może być pusta")
            .Must(x => x == null || x.Length <= 100 * 1024 * 1024).WithMessage("Rozmiar dokumentu nie może przekraczać 100 MB");

        RuleFor(x => x.CreatedBy)
            .NotEmpty().WithMessage("Pole 'CreatedBy' jest wymagane")
            .MaximumLength(255).WithMessage("Pole 'CreatedBy' nie może przekraczać 255 znaków");

        RuleFor(x => x.Metadata)
            .MaximumLength(8000).WithMessage("Metadata nie może przekraczać 8000 znaków")
            .Must(BeValidJson).WithMessage("Metadata musi być poprawnym JSON-em")
            .When(x => !string.IsNullOrEmpty(x.Metadata));
    }

    private static bool BeSupportedMimeType(string mimeType)
        => string.Equals(mimeType, IngestExternalDocumentCommandHandler.DocxMimeType, StringComparison.OrdinalIgnoreCase)
        || string.Equals(mimeType, IngestExternalDocumentCommandHandler.PdfMimeType, StringComparison.OrdinalIgnoreCase);

    private static bool BeValidJson(string? metadata)
    {
        if (string.IsNullOrWhiteSpace(metadata))
            return true;

        try
        {
            using var _ = System.Text.Json.JsonDocument.Parse(metadata);
            return true;
        }
        catch (System.Text.Json.JsonException)
        {
            return false;
        }
    }
}
