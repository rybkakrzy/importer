using Microsoft.AspNetCore.Mvc;
using ImporterApi.Models;
using ImporterApi.Services;

namespace ImporterApi.Controllers;

/// <summary>
/// Kontroler API do generowania kodów kreskowych i QR
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class BarcodeController : ControllerBase
{
    private readonly BarcodeGeneratorService _barcodeService;

    public BarcodeController(BarcodeGeneratorService barcodeService)
    {
        _barcodeService = barcodeService;
    }

    /// <summary>
    /// Pobiera listę obsługiwanych typów kodów
    /// </summary>
    [HttpGet("types")]
    public ActionResult<IEnumerable<string>> GetSupportedTypes()
    {
        return Ok(BarcodeGeneratorService.GetSupportedTypes());
    }

    /// <summary>
    /// Generuje kod kreskowy / QR i zwraca jako base64
    /// </summary>
    [HttpPost("generate")]
    public ActionResult<BarcodeResponse> GenerateBarcode([FromBody] BarcodeRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.Content))
        {
            return BadRequest(new { error = "Treść kodu nie może być pusta" });
        }

        if (string.IsNullOrWhiteSpace(request.BarcodeType))
        {
            return BadRequest(new { error = "Typ kodu jest wymagany" });
        }

        // Walidacja wymiarów
        var width = Math.Clamp(request.Width, 50, 1000);
        var height = Math.Clamp(request.Height, 50, 1000);

        try
        {
            var pngBytes = _barcodeService.GenerateBarcode(
                request.Content,
                request.BarcodeType,
                width,
                height
            );

            var base64 = Convert.ToBase64String(pngBytes);

            return Ok(new BarcodeResponse
            {
                Base64Image = $"data:image/png;base64,{base64}",
                ContentType = "image/png",
                BarcodeType = request.BarcodeType
            });
        }
        catch (ArgumentException ex)
        {
            return BadRequest(new { error = ex.Message });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = $"Błąd generowania kodu: {ex.Message}" });
        }
    }

    /// <summary>
    /// Generuje kod kreskowy / QR i zwraca jako obraz PNG
    /// </summary>
    [HttpPost("generate-image")]
    public ActionResult GenerateBarcodeImage([FromBody] BarcodeRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.Content))
        {
            return BadRequest(new { error = "Treść kodu nie może być pusta" });
        }

        var width = Math.Clamp(request.Width, 50, 1000);
        var height = Math.Clamp(request.Height, 50, 1000);

        try
        {
            var pngBytes = _barcodeService.GenerateBarcode(
                request.Content,
                request.BarcodeType,
                width,
                height
            );

            return File(pngBytes, "image/png", $"{request.BarcodeType}.png");
        }
        catch (ArgumentException ex)
        {
            return BadRequest(new { error = ex.Message });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = $"Błąd generowania kodu: {ex.Message}" });
        }
    }
}
