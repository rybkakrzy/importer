using FluentValidation;
using Importer.Application.Features.Documents.Commands.SaveDocument;

namespace Importer.Application.Validators;

public class SaveDocumentValidator : AbstractValidator<SaveDocumentCommand>
{
    public SaveDocumentValidator()
    {
        RuleFor(x => x.Html)
            .NotEmpty().WithMessage("Treść HTML dokumentu jest wymagana");
    }
}
