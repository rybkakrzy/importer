using FluentValidation;
using D2ViewerEditor.Application.Features.Documents.Commands.SaveDocument;

namespace D2ViewerEditor.Application.Validators;

public class SaveDocumentValidator : AbstractValidator<SaveDocumentCommand>
{
    public SaveDocumentValidator()
    {
        RuleFor(x => x.Html)
            .NotEmpty().WithMessage("Treść HTML dokumentu jest wymagana");
    }
}
