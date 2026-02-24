using D2Tools.Application.Features.Barcode.Commands.GenerateBarcode;
using D2Tools.Application.Features.Barcode.Queries.GetSupportedTypes;
using D2Tools.Domain.Models;
using Microsoft.AspNetCore.Mvc;

namespace D2Tools.Api.Controllers;

/// <summary>
/// Kontroler API do generowania kodów kreskowych i QR
/// </summary>
[Route("api/[controller]")]
public class BarcodeController : BaseApiController
{
    /// <summary>
    /// Pobiera listę obsługiwanych typów kodów
    /// </summary>
    [HttpGet("types")]
    public async Task<IActionResult> GetSupportedTypes()
    {
        var result = await Mediator.Send(new GetSupportedTypesQuery());
        return HandleResult(result);
    }

    /// <summary>
    /// Generuje kod kreskowy / QR i zwraca jako base64
    /// </summary>
    [HttpPost("generate")]
    public async Task<IActionResult> GenerateBarcode([FromBody] BarcodeRequest request)
    {
        var command = new GenerateBarcodeCommand(
            request.Content,
            request.BarcodeType,
            request.Width,
            request.Height,
            request.ShowText
        );

        var result = await Mediator.Send(command);
        return HandleResult(result);
    }

    /// <summary>
    /// Generuje kod kreskowy / QR i zwraca jako obraz PNG
    /// </summary>
    [HttpPost("generate-image")]
    public async Task<IActionResult> GenerateBarcodeImage([FromBody] BarcodeRequest request)
    {
        var command = new GenerateBarcodeCommand(
            request.Content,
            request.BarcodeType,
            request.Width,
            request.Height,
            request.ShowText
        );

        var result = await Mediator.Send(command);

        return result.Match<IActionResult>(
            onSuccess: response =>
            {
                var base64Data = response.Base64Image.Replace("data:image/png;base64,", "");
                var bytes = Convert.FromBase64String(base64Data);
                return File(bytes, "image/png", $"{request.BarcodeType}.png");
            },
            onFailure: error => BadRequest(new { error })
        );
    }
}
