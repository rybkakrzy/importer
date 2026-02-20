using FluentValidation;
using Importer.Application.Features.Barcode.Commands.GenerateBarcode;

namespace Importer.Application.Validators;

public class GenerateBarcodeValidator : AbstractValidator<GenerateBarcodeCommand>
{
    public GenerateBarcodeValidator()
    {
        RuleFor(x => x.Content)
            .NotEmpty().WithMessage("Treść kodu kreskowego jest wymagana")
            .MaximumLength(500).WithMessage("Treść nie może przekraczać 500 znaków");

        RuleFor(x => x.BarcodeType)
            .NotEmpty().WithMessage("Typ kodu kreskowego jest wymagany");

        RuleFor(x => x.Width)
            .InclusiveBetween(50, 2000).WithMessage("Szerokość musi być między 50 a 2000 px");

        RuleFor(x => x.Height)
            .InclusiveBetween(50, 2000).WithMessage("Wysokość musi być między 50 a 2000 px");
    }
}
