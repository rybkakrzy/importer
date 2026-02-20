using FluentValidation;
using Importer.Application.Features.FileUpload.Commands.ProcessZipFile;

namespace Importer.Application.Validators;

public class ProcessZipFileValidator : AbstractValidator<ProcessZipFileCommand>
{
    public ProcessZipFileValidator()
    {
        RuleFor(x => x.FileName)
            .NotEmpty().WithMessage("Nazwa pliku jest wymagana")
            .Must(x => x.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
            .WithMessage("Plik musi mieć rozszerzenie .zip");
    }
}
