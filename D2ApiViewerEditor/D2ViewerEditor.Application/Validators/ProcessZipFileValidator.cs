using FluentValidation;
using D2ViewerEditor.Application.Features.FileUpload.Commands.ProcessZipFile;

namespace D2ViewerEditor.Application.Validators;

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
