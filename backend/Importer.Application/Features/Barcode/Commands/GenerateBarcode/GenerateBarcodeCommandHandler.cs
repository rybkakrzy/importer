using Importer.Domain.Common;
using Importer.Domain.Interfaces;
using Importer.Domain.Models;
using MediatR;

namespace Importer.Application.Features.Barcode.Commands.GenerateBarcode;

public class GenerateBarcodeCommandHandler : IRequestHandler<GenerateBarcodeCommand, Result<BarcodeResponse>>
{
    private readonly IBarcodeGenerator _barcodeGenerator;

    public GenerateBarcodeCommandHandler(IBarcodeGenerator barcodeGenerator)
    {
        _barcodeGenerator = barcodeGenerator;
    }

    public Task<Result<BarcodeResponse>> Handle(GenerateBarcodeCommand request, CancellationToken cancellationToken)
    {
        var width = Math.Clamp(request.Width, 50, 1000);
        var height = Math.Clamp(request.Height, 50, 1000);

        try
        {
            var pngBytes = _barcodeGenerator.Generate(
                request.Content, request.BarcodeType, width, height, request.ShowText);

            var base64 = Convert.ToBase64String(pngBytes);

            var response = new BarcodeResponse
            {
                Base64Image = $"data:image/png;base64,{base64}",
                ContentType = "image/png",
                BarcodeType = request.BarcodeType
            };

            return Task.FromResult(Result<BarcodeResponse>.Success(response));
        }
        catch (ArgumentException ex)
        {
            return Task.FromResult(Result<BarcodeResponse>.Failure(ex.Message));
        }
    }
}
