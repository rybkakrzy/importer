using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Documents.Commands.SaveDocument;

public class SaveDocumentCommandHandler : IRequestHandler<SaveDocumentCommand, Result<SaveDocumentResult>>
{
    private readonly IHtmlToDocxConverter _converter;

    public SaveDocumentCommandHandler(IHtmlToDocxConverter converter)
    {
        _converter = converter;
    }

    public Task<Result<SaveDocumentResult>> Handle(SaveDocumentCommand request, CancellationToken cancellationToken)
    {
        var docxBytes = _converter.Convert(request.Html, request.Metadata, request.Header, request.Footer, request.Margins, request.PageSize, request.SectionHeadersFooters);

        var fileName = !string.IsNullOrEmpty(request.OriginalFileName)
            ? request.OriginalFileName
            : "dokument.docx";

        if (!fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            fileName += ".docx";

        return Task.FromResult(Result<SaveDocumentResult>.Success(new SaveDocumentResult(docxBytes, fileName)));
    }
}
