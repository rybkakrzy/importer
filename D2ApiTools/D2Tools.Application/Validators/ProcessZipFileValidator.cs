using FluentValidation;
using D2Tools.Application.Features.FileUpload.Commands.ProcessZipFile;

namespace D2Tools.Application.Validators;

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
