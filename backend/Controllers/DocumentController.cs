using Microsoft.AspNetCore.Mvc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace ImporterApi.Controllers;

[ApiController]
[Route("api/[controller]")]
public class DocumentController : ControllerBase
{
    private readonly ILogger<DocumentController> _logger;

    public DocumentController(ILogger<DocumentController> logger)
    {
        _logger = logger;
    }

    [HttpPost("upload")]
    [Consumes("application/octet-stream")]
    public async Task<ActionResult<DocumentUploadResponse>> UploadDocument()
    {
        var fileName = Request.Headers["X-File-Name"].ToString();
        fileName = Uri.UnescapeDataString(fileName);

        _logger.LogInformation($"Uploading document: {fileName}");

        if (string.IsNullOrEmpty(fileName) || !fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
        {
            return BadRequest(new DocumentUploadResponse
            {
                Success = false,
                Message = "File must be a DOCX document"
            });
        }

        try
        {
            using var memoryStream = new MemoryStream();
            await Request.Body.CopyToAsync(memoryStream);

            if (memoryStream.Length == 0)
            {
                return BadRequest(new DocumentUploadResponse
                {
                    Success = false,
                    Message = "File is empty"
                });
            }

            memoryStream.Position = 0;
            var content = ExtractDocumentContent(memoryStream);

            return Ok(new DocumentUploadResponse
            {
                Success = true,
                Message = "Document processed successfully",
                FileName = fileName,
                Content = content
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing document");
            return StatusCode(500, new DocumentUploadResponse
            {
                Success = false,
                Message = "Error processing document",
                ErrorDetails = ex.Message
            });
        }
    }

    [HttpPost("save")]
    public async Task<ActionResult> SaveDocument([FromBody] DocumentSaveRequest request)
    {
        if (string.IsNullOrEmpty(request.Content))
        {
            return BadRequest("Content cannot be empty");
        }

        try
        {
            var memoryStream = new MemoryStream();
            CreateDocumentFromContent(memoryStream, request.Content);
            
            memoryStream.Position = 0;
            var fileName = request.FileName ?? "document.docx";
            
            return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error saving document");
            return StatusCode(500, "Error saving document");
        }
    }

    private string ExtractDocumentContent(Stream stream)
    {
        var content = new List<DocumentParagraph>();

        using (var wordDoc = WordprocessingDocument.Open(stream, false))
        {
            var body = wordDoc.MainDocumentPart?.Document?.Body;
            if (body == null)
                return string.Empty;

            foreach (var element in body.Elements())
            {
                if (element is Paragraph paragraph)
                {
                    var paragraphContent = ExtractParagraphContent(paragraph);
                    if (!string.IsNullOrEmpty(paragraphContent.Text))
                    {
                        content.Add(paragraphContent);
                    }
                }
            }
        }

        return System.Text.Json.JsonSerializer.Serialize(content);
    }

    private DocumentParagraph ExtractParagraphContent(Paragraph paragraph)
    {
        var runs = new List<DocumentRun>();
        var paragraphProperties = paragraph.ParagraphProperties;
        var style = "normal";

        if (paragraphProperties?.ParagraphStyleId != null)
        {
            var styleId = paragraphProperties.ParagraphStyleId.Val?.Value ?? "";
            if (styleId.Contains("Heading1")) style = "h1";
            else if (styleId.Contains("Heading2")) style = "h2";
            else if (styleId.Contains("Heading3")) style = "h3";
        }

        foreach (var run in paragraph.Elements<Run>())
        {
            var text = run.InnerText;
            if (string.IsNullOrEmpty(text))
                continue;

            var runProperties = run.RunProperties;
            var isBold = runProperties?.Bold != null;
            var isItalic = runProperties?.Italic != null;
            var isUnderline = runProperties?.Underline != null;

            runs.Add(new DocumentRun
            {
                Text = text,
                Bold = isBold,
                Italic = isItalic,
                Underline = isUnderline
            });
        }

        return new DocumentParagraph
        {
            Text = paragraph.InnerText,
            Style = style,
            Runs = runs
        };
    }

    private void CreateDocumentFromContent(Stream stream, string jsonContent)
    {
        var paragraphs = System.Text.Json.JsonSerializer.Deserialize<List<DocumentParagraph>>(jsonContent);
        if (paragraphs == null)
            throw new InvalidOperationException("Invalid content format");

        using (var wordDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            foreach (var para in paragraphs)
            {
                var paragraph = new Paragraph();

                // Add style if not normal
                if (para.Style != "normal")
                {
                    var paragraphProperties = new ParagraphProperties();
                    var paragraphStyleId = new ParagraphStyleId { Val = para.Style switch
                    {
                        "h1" => "Heading1",
                        "h2" => "Heading2",
                        "h3" => "Heading3",
                        _ => "Normal"
                    }};
                    paragraphProperties.Append(paragraphStyleId);
                    paragraph.Append(paragraphProperties);
                }

                // Add runs with formatting
                if (para.Runs != null && para.Runs.Count > 0)
                {
                    foreach (var runData in para.Runs)
                    {
                        var run = new Run();
                        var runProperties = new RunProperties();

                        if (runData.Bold)
                            runProperties.Append(new Bold());
                        if (runData.Italic)
                            runProperties.Append(new Italic());
                        if (runData.Underline)
                            runProperties.Append(new Underline { Val = UnderlineValues.Single });

                        if (runProperties.HasChildren)
                            run.Append(runProperties);

                        run.Append(new Text(runData.Text));
                        paragraph.Append(run);
                    }
                }
                else
                {
                    // Fallback: create a simple run with the paragraph text
                    var run = new Run(new Text(para.Text ?? ""));
                    paragraph.Append(run);
                }

                body.Append(paragraph);
            }

            mainPart.Document.Save();
        }
    }
}

public class DocumentUploadResponse
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public string? FileName { get; set; }
    public string? Content { get; set; }
    public string? ErrorDetails { get; set; }
}

public class DocumentSaveRequest
{
    public string Content { get; set; } = string.Empty;
    public string? FileName { get; set; }
}

public class DocumentParagraph
{
    public string Text { get; set; } = string.Empty;
    public string Style { get; set; } = "normal";
    public List<DocumentRun> Runs { get; set; } = new();
}

public class DocumentRun
{
    public string Text { get; set; } = string.Empty;
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
}
