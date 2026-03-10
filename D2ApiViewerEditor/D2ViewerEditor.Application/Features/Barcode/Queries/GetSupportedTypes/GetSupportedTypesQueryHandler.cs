using D2ViewerEditor.Domain.Common;
using D2ViewerEditor.Domain.Interfaces;
using MediatR;

namespace D2ViewerEditor.Application.Features.Barcode.Queries.GetSupportedTypes;

public class GetSupportedTypesQueryHandler : IRequestHandler<GetSupportedTypesQuery, Result<IReadOnlyList<string>>>
{
    private readonly IBarcodeGenerator _barcodeGenerator;

    public GetSupportedTypesQueryHandler(IBarcodeGenerator barcodeGenerator)
    {
        _barcodeGenerator = barcodeGenerator;
    }

    public Task<Result<IReadOnlyList<string>>> Handle(GetSupportedTypesQuery request, CancellationToken cancellationToken)
    {
        var types = _barcodeGenerator.GetSupportedTypes();
        return Task.FromResult(Result<IReadOnlyList<string>>.Success(types));
    }
}
