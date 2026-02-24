using FluentValidation;
using D2Tools.Application.Features.Documents.Commands.UploadImage;

namespace D2Tools.Application.Validators;

public class UploadImageValidator : AbstractValidator<UploadImageCommand>
{
    private const long MaxFileSize = 10 * 1024 * 1024; // 10 MB

    public UploadImageValidator()
    {
        RuleFor(x => x.FileName)
            .NotEmpty().WithMessage("Nazwa pliku jest wymagana");

        RuleFor(x => x.ContentType)
            .NotEmpty().WithMessage("Typ pliku jest wymagany")
            .Must(BeAnImageType).WithMessage("Plik musi być obrazem (image/*)");

        RuleFor(x => x.FileSize)
            .LessThanOrEqualTo(MaxFileSize).WithMessage("Rozmiar obrazu nie może przekraczać 10 MB");
    }

    private static bool BeAnImageType(string contentType)
        => contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase);
}
