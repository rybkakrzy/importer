using FluentValidation;
using D2Tools.Application.Features.Documents.Commands.SaveDocument;

namespace D2Tools.Application.Validators;

public class SaveDocumentValidator : AbstractValidator<SaveDocumentCommand>
{
    public SaveDocumentValidator()
    {
        RuleFor(x => x.Html)
            .NotEmpty().WithMessage("Treść HTML dokumentu jest wymagana");
    }
}
