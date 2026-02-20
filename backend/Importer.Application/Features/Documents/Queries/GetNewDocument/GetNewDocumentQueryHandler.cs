using Importer.Domain.Common;
using Importer.Domain.Models;
using MediatR;

namespace Importer.Application.Features.Documents.Queries.GetNewDocument;

public class GetNewDocumentQueryHandler : IRequestHandler<GetNewDocumentQuery, Result<DocumentContent>>
{
    public Task<Result<DocumentContent>> Handle(GetNewDocumentQuery request, CancellationToken cancellationToken)
    {
        var content = new DocumentContent
        {
            Html = "<div class=\"document-content\"><p>&nbsp;</p></div>",
            Metadata = new DocumentMetadata
            {
                Title = "Nowy dokument",
                Created = DateTime.Now,
                Modified = DateTime.Now
            },
            Images = new List<DocumentImage>()
        };

        return Task.FromResult(Result<DocumentContent>.Success(content));
    }
}
