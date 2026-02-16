using Microsoft.AspNetCore.Mvc;
using ImporterApi.Models;
using ImporterApi.Services;

namespace ImporterApi.Controllers;

/// <summary>
/// Kontroler API dla edytora dokumentów Word
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class DocumentController : ControllerBase
{
    private readonly DocxToHtmlConverter _docxToHtml;
    private readonly HtmlToDocxConverter _htmlToDocx;
    private readonly IWebHostEnvironment _environment;
    private readonly string _uploadsPath;

    public DocumentController(
        DocxToHtmlConverter docxToHtml,
        HtmlToDocxConverter htmlToDocx,
        IWebHostEnvironment environment)
    {
        _docxToHtml = docxToHtml;
        _htmlToDocx = htmlToDocx;
        _environment = environment;
        _uploadsPath = Path.Combine(_environment.ContentRootPath, "Uploads");
        
        if (!Directory.Exists(_uploadsPath))
        {
            Directory.CreateDirectory(_uploadsPath);
        }
    }

    /// <summary>
    /// Wczytuje dokument DOCX i zwraca HTML
    /// </summary>
    [HttpPost("open")]
    [RequestSizeLimit(50 * 1024 * 1024)] // 50 MB limit
    public async Task<ActionResult<DocumentContent>> OpenDocument(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest(new { error = "Nie przesłano pliku" });
        }

        var extension = Path.GetExtension(file.FileName).ToLower();
        if (extension != ".docx")
        {
            return BadRequest(new { error = "Obsługiwane są tylko pliki DOCX" });
        }

        try
        {
            using var stream = file.OpenReadStream();
            using var memoryStream = new MemoryStream();
            await stream.CopyToAsync(memoryStream);
            memoryStream.Position = 0;

            var content = _docxToHtml.Convert(memoryStream);
            content.Metadata.Title ??= Path.GetFileNameWithoutExtension(file.FileName);

            return Ok(content);
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = $"Błąd podczas przetwarzania dokumentu: {ex.Message}" });
        }
    }

    /// <summary>
    /// Zapisuje HTML jako dokument DOCX
    /// </summary>
    [HttpPost("save")]
    public ActionResult SaveDocument([FromBody] SaveDocumentRequest request)
    {
        if (string.IsNullOrEmpty(request.Html))
        {
            return BadRequest(new { error = "Brak zawartości dokumentu" });
        }

        try
        {
            var docxBytes = _htmlToDocx.Convert(request.Html, request.Metadata, request.Header, request.Footer);
            
            var fileName = !string.IsNullOrEmpty(request.OriginalFileName) 
                ? request.OriginalFileName 
                : "dokument.docx";
            
            if (!fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                fileName += ".docx";
            }

            return File(docxBytes, 
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                fileName);
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = $"Błąd podczas zapisywania dokumentu: {ex.Message}" });
        }
    }

    /// <summary>
    /// Tworzy nowy pusty dokument
    /// </summary>
    [HttpGet("new")]
    public ActionResult<DocumentContent> NewDocument()
    {
        return Ok(new DocumentContent
        {
            Html = "<div class=\"document-content\"><p>&nbsp;</p></div>",
            Metadata = new DocumentMetadata
            {
                Title = "Nowy dokument",
                Created = DateTime.Now,
                Modified = DateTime.Now
            },
            Images = new List<DocumentImage>()
        });
    }

    /// <summary>
    /// Eksportuje dokument do PDF (konwersja przez DOCX)
    /// </summary>
    [HttpPost("export-pdf")]
    public ActionResult ExportToPdf([FromBody] SaveDocumentRequest request)
    {
        // Uwaga: Export do PDF wymaga dodatkowej biblioteki lub zewnętrznego serwisu
        // Na razie zwracamy informację że funkcja nie jest dostępna
        return StatusCode(501, new { error = "Eksport do PDF nie jest jeszcze zaimplementowany. Użyj funkcji Zapisz jako DOCX." });
    }

    /// <summary>
    /// Wgrywa obraz i zwraca Base64
    /// </summary>
    [HttpPost("upload-image")]
    [RequestSizeLimit(10 * 1024 * 1024)] // 10 MB limit dla obrazów
    public async Task<ActionResult> UploadImage(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest(new { error = "Nie przesłano pliku" });
        }

        var allowedExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp" };
        var extension = Path.GetExtension(file.FileName).ToLower();
        
        if (!allowedExtensions.Contains(extension))
        {
            return BadRequest(new { error = "Niedozwolony format obrazu" });
        }

        try
        {
            using var memoryStream = new MemoryStream();
            await file.CopyToAsync(memoryStream);
            var base64 = Convert.ToBase64String(memoryStream.ToArray());
            var contentType = file.ContentType;

            return Ok(new
            {
                base64 = $"data:{contentType};base64,{base64}",
                fileName = file.FileName,
                size = file.Length
            });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = $"Błąd podczas wgrywania obrazu: {ex.Message}" });
        }
    }

    /// <summary>
    /// Pobiera listę szablonów dokumentów
    /// </summary>
    [HttpGet("templates")]
    public ActionResult<List<object>> GetTemplates()
    {
        var templates = new List<object>
        {
            new { id = "blank", name = "Pusty dokument", description = "Zacznij od pustej strony" },
            new { id = "letter", name = "List", description = "Szablon listu formalnego" },
            new { id = "report", name = "Raport", description = "Szablon raportu z nagłówkami" },
            new { id = "cv", name = "CV", description = "Szablon curriculum vitae" }
        };

        return Ok(templates);
    }

    /// <summary>
    /// Pobiera szablon dokumentu
    /// </summary>
    [HttpGet("templates/{templateId}")]
    public ActionResult<DocumentContent> GetTemplate(string templateId)
    {
        var html = templateId switch
        {
            "blank" => "<div class=\"document-content\"><p>&nbsp;</p></div>",
            
            "letter" => @"<div class=""document-content"">
                <p style=""text-align:right;"">Miejscowość, data</p>
                <p>&nbsp;</p>
                <p><strong>Nadawca:</strong></p>
                <p>Imię i Nazwisko</p>
                <p>Adres</p>
                <p>&nbsp;</p>
                <p><strong>Odbiorca:</strong></p>
                <p>Imię i Nazwisko</p>
                <p>Adres</p>
                <p>&nbsp;</p>
                <p style=""text-align:center;""><strong>Temat listu</strong></p>
                <p>&nbsp;</p>
                <p>Szanowni Państwo,</p>
                <p>&nbsp;</p>
                <p>Treść listu...</p>
                <p>&nbsp;</p>
                <p>Z poważaniem,</p>
                <p>Podpis</p>
            </div>",
            
            "report" => @"<div class=""document-content"">
                <h1 style=""text-align:center;"">Tytuł Raportu</h1>
                <p style=""text-align:center;"">Autor: Imię Nazwisko</p>
                <p style=""text-align:center;"">Data: " + DateTime.Now.ToString("dd.MM.yyyy") + @"</p>
                <p>&nbsp;</p>
                <h2>1. Wprowadzenie</h2>
                <p>Treść wprowadzenia...</p>
                <p>&nbsp;</p>
                <h2>2. Analiza</h2>
                <p>Treść analizy...</p>
                <p>&nbsp;</p>
                <h2>3. Wnioski</h2>
                <p>Treść wniosków...</p>
                <p>&nbsp;</p>
                <h2>4. Podsumowanie</h2>
                <p>Treść podsumowania...</p>
            </div>",
            
            "cv" => @"<div class=""document-content"">
                <h1 style=""text-align:center;"">Imię i Nazwisko</h1>
                <p style=""text-align:center;"">email@example.com | +48 123 456 789 | Miasto</p>
                <p>&nbsp;</p>
                <h2>Doświadczenie zawodowe</h2>
                <p><strong>Stanowisko</strong> | Firma | Okres</p>
                <p>Opis obowiązków...</p>
                <p>&nbsp;</p>
                <h2>Wykształcenie</h2>
                <p><strong>Kierunek</strong> | Uczelnia | Okres</p>
                <p>&nbsp;</p>
                <h2>Umiejętności</h2>
                <p>• Umiejętność 1</p>
                <p>• Umiejętność 2</p>
                <p>• Umiejętność 3</p>
                <p>&nbsp;</p>
                <h2>Języki</h2>
                <p>• Polski - ojczysty</p>
                <p>• Angielski - zaawansowany</p>
            </div>",
            
            _ => "<div class=\"document-content\"><p>&nbsp;</p></div>"
        };

        var title = templateId switch
        {
            "letter" => "Nowy list",
            "report" => "Nowy raport",
            "cv" => "Nowe CV",
            _ => "Nowy dokument"
        };

        return Ok(new DocumentContent
        {
            Html = html,
            Metadata = new DocumentMetadata
            {
                Title = title,
                Created = DateTime.Now,
                Modified = DateTime.Now
            },
            Images = new List<DocumentImage>()
        });
    }
}
