using Importer.Domain.Common;
using Importer.Domain.Interfaces;
using Importer.Domain.Models;
using MediatR;

namespace Importer.Application.Features.Documents.Queries.OpenDocument;

public class OpenDocumentQueryHandler : IRequestHandler<OpenDocumentQuery, Result<DocumentContent>>
{
    private readonly IDocxToHtmlConverter _converter;

    public OpenDocumentQueryHandler(IDocxToHtmlConverter converter)
    {
        _converter = converter;
    }

    public async Task<Result<DocumentContent>> Handle(OpenDocumentQuery request, CancellationToken cancellationToken)
    {
        using var memoryStream = new MemoryStream();
        await request.FileStream.CopyToAsync(memoryStream, cancellationToken);
        memoryStream.Position = 0;

        var content = _converter.Convert(memoryStream);
        content.Metadata.Title ??= Path.GetFileNameWithoutExtension(request.FileName);

        return Result<DocumentContent>.Success(content);
    }
}
