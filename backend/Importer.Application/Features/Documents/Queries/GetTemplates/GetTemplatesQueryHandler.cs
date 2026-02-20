using Importer.Domain.Common;
using MediatR;

namespace Importer.Application.Features.Documents.Queries.GetTemplates;

public class GetTemplatesQueryHandler : IRequestHandler<GetTemplatesQuery, Result<List<TemplateDto>>>
{
    public Task<Result<List<TemplateDto>>> Handle(GetTemplatesQuery request, CancellationToken cancellationToken)
    {
        var templates = new List<TemplateDto>
        {
            new("blank", "Pusty dokument", "Zacznij od pustej strony"),
            new("letter", "List", "Szablon listu formalnego"),
            new("report", "Raport", "Szablon raportu z nagłówkami"),
            new("cv", "CV", "Szablon curriculum vitae")
        };

        return Task.FromResult(Result<List<TemplateDto>>.Success(templates));
    }
}
