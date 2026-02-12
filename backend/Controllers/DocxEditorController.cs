using System.IO.Compression;
using System.Net;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.AspNetCore.Mvc;

namespace ImporterApi.Controllers;

[ApiController]
[Route("api/[controller]")]
public class DocxEditorController : ControllerBase
{
    private const string DocxContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    private const long MaxDocxSizeBytes = 10 * 1024 * 1024;

    [HttpPost("open")]
    [Consumes("application/octet-stream")]
    public async Task<ActionResult<DocxOpenResponse>> OpenDocx()
    {
        var fileName = Uri.UnescapeDataString(Request.Headers["X-File-Name"].ToString());
        if (string.IsNullOrWhiteSpace(fileName) || !fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
        {
            return BadRequest(new DocxOpenResponse { Success = false, Message = "Przesłany plik musi być typu DOCX" });
        }
        var safeFileName = $"{SanitizeFileName(fileName)}.docx";

        await using var memoryStream = new MemoryStream();
        await Request.Body.CopyToAsync(memoryStream);
        if (memoryStream.Length == 0 || memoryStream.Length > MaxDocxSizeBytes)
        {
            return BadRequest(new DocxOpenResponse { Success = false, Message = "Nieprawidłowy rozmiar pliku DOCX" });
        }

        memoryStream.Position = 0;
        try
        {
            using var archive = new ZipArchive(memoryStream, ZipArchiveMode.Read, leaveOpen: false);
            var documentEntry = archive.GetEntry("word/document.xml");
            if (documentEntry is null)
            {
                return BadRequest(new DocxOpenResponse { Success = false, Message = "Brak dokumentu word/document.xml w pliku DOCX" });
            }

            using var documentStream = documentEntry.Open();
            var xDocument = XDocument.Load(documentStream);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var paragraphs = xDocument
                .Descendants(w + "p")
                .Select(p => string.Concat(p.Descendants(w + "t").Select(t => (string?)t ?? string.Empty)))
                .ToList();

            if (paragraphs.Count == 0)
            {
                paragraphs.Add(string.Concat(xDocument.Descendants(w + "t").Select(t => (string?)t ?? string.Empty)));
            }

            var html = string.Join(string.Empty, paragraphs.Select(p => $"<p>{WebUtility.HtmlEncode(p)}</p>"));
            if (string.IsNullOrEmpty(html))
            {
                html = "<p><br></p>";
            }

            return Ok(new DocxOpenResponse
            {
                Success = true,
                Message = "Plik DOCX został odczytany",
                FileName = safeFileName,
                Html = html
            });
        }
        catch (InvalidDataException)
        {
            return BadRequest(new DocxOpenResponse { Success = false, Message = "Nieprawidłowy format pliku DOCX" });
        }
    }

    [HttpPost("save")]
    public ActionResult SaveDocx([FromBody] DocxSaveRequest request)
    {
        if (!HasMeaningfulText(request.Html))
        {
            return BadRequest("Brak treści do zapisu.");
        }

        var paragraphs = ExtractParagraphsFromHtml(request.Html);
        var safeFileName = SanitizeFileName(request.FileName);
        var documentXml = BuildDocumentXml(paragraphs);

        using var memoryStream = new MemoryStream();
        using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "[Content_Types].xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """);

            WriteEntry(archive, "_rels/.rels",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """);

            WriteEntry(archive, "word/document.xml", documentXml);
        }

        return File(memoryStream.ToArray(), DocxContentType, $"{safeFileName}.docx");
    }

    private static string BuildDocumentXml(IEnumerable<string> paragraphs)
    {
        var body = new StringBuilder();
        foreach (var paragraph in paragraphs)
        {
            if (string.IsNullOrEmpty(paragraph))
            {
                body.Append("<w:p/>");
                continue;
            }

            body.Append("<w:p><w:r><w:t xml:space=\"preserve\">")
                .Append(SecurityElement.Escape(paragraph))
                .Append("</w:t></w:r></w:p>");
        }

        return $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    {body}
                  </w:body>
                </w:document>
                """;
    }

    private static List<string> ExtractParagraphsFromHtml(string html)
    {
        var noScripts = Regex.Replace(html, @"<(script|style)[^>]*>.*?</\1>", string.Empty, RegexOptions.Singleline | RegexOptions.IgnoreCase);
        var withBreaks = Regex.Replace(noScripts, @"<br\s*/?>", "\n", RegexOptions.IgnoreCase);
        var withParagraphBreaks = Regex.Replace(withBreaks, @"</(p|div|li|h[1-6])>", "\n", RegexOptions.IgnoreCase);
        var noTags = Regex.Replace(withParagraphBreaks, @"<[^>]+>", string.Empty);
        var decoded = WebUtility.HtmlDecode(noTags);

        var paragraphs = decoded
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Split('\n')
            .Select(p => p.Trim())
            .ToList();

        return paragraphs.All(string.IsNullOrWhiteSpace) ? [string.Empty] : paragraphs;
    }

    private static void WriteEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(content);
    }

    private static string SanitizeFileName(string? fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return "dokument";
        }

        var noExtension = Path.GetFileNameWithoutExtension(fileName);
        var sanitized = string.Concat(noExtension.Where(ch => !Path.GetInvalidFileNameChars().Contains(ch)));
        return string.IsNullOrWhiteSpace(sanitized) ? "dokument" : sanitized;
    }

    private static bool HasMeaningfulText(string? html)
    {
        if (string.IsNullOrWhiteSpace(html))
        {
            return false;
        }

        var text = WebUtility.HtmlDecode(Regex.Replace(html, "<[^>]+>", string.Empty));
        return !string.IsNullOrWhiteSpace(text);
    }
}

public class DocxOpenResponse
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public string? FileName { get; set; }
    public string? Html { get; set; }
}

public class DocxSaveRequest
{
    public string? FileName { get; set; }
    public string Html { get; set; } = string.Empty;
}
