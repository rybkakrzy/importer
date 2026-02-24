using D2Tools.Application.Features.FileUpload.Commands.ProcessZipFile;
using Microsoft.AspNetCore.Mvc;

namespace D2Tools.Api.Controllers;

/// <summary>
/// Kontroler API do przetwarzania plików ZIP
/// </summary>
[Route("api/[controller]")]
public class FileUploadController : BaseApiController
{
    [HttpPost("upload")]
    [Consumes("application/octet-stream")]
    public async Task<IActionResult> UploadZipFile()
    {
        var fileName = Request.Headers["X-File-Name"].ToString();
        fileName = Uri.UnescapeDataString(fileName);

        if (string.IsNullOrEmpty(fileName) || !fileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
        {
            return BadRequest(new { error = "Przesłany plik musi być archiwum ZIP" });
        }

        var command = new ProcessZipFileCommand(Request.Body, fileName);
        var result = await Mediator.Send(command);

        return result.Match<IActionResult>(
            onSuccess: zipResult => Ok(new
            {
                success = true,
                message = "Plik ZIP został przetworzony pomyślnie",
                fileName,
                fileSize = zipResult.TotalSize,
                filesCount = zipResult.TotalFiles,
                files = zipResult.Files.Select(f => new
                {
                    fileName = f.FileName,
                    fileSize = f.Size,
                    content = f.Content
                }),
                processedAt = DateTime.UtcNow
            }),
            onFailure: error => BadRequest(new { success = false, message = error })
        );
    }
}
