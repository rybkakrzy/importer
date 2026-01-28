using Microsoft.AspNetCore.Mvc;
using System.IO.Compression;

namespace ImporterApi.Controllers;

[ApiController]
[Route("api/[controller]")]
public class FileUploadController : ControllerBase
{
    private readonly ILogger<FileUploadController> _logger;

    public FileUploadController(ILogger<FileUploadController> logger)
    {
        _logger = logger;
    }

    [HttpPost("upload")]
    [Consumes("application/octet-stream")]
    public async Task<ActionResult<ZipUploadResponse>> UploadZipFile()
    {
        // Pobierz nazwę pliku z nagłówka
        var fileName = Request.Headers["X-File-Name"].ToString();
        fileName = Uri.UnescapeDataString(fileName);
        
        _logger.LogTrace($"[Walidator_Byte] -> Wczytano plik: '{fileName}'");

        if (string.IsNullOrEmpty(fileName) || !fileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
        {
            return BadRequest(new ZipUploadResponse
            {
                Success = false,
                Message = "Przesłany plik musi być archiwum ZIP",
                ErrorCode = "INVALID_FILE_TYPE"
            });
        }

        try
        {
            using var memoryStream = new MemoryStream();
            await Request.Body.CopyToAsync(memoryStream);
            
            if (memoryStream.Length == 0)
            {
                return BadRequest(new ZipUploadResponse
                {
                    Success = false,
                    Message = "Plik jest pusty",
                    ErrorCode = "EMPTY_FILE"
                });
            }

            _logger.LogTrace($"[Walidator_Byte] -> Plik wysłany przez: '{User.Identity?.Name ?? "Nieznany użytkownik"}'");
            
            memoryStream.Position = 0;
            var fileSize = memoryStream.Length;
            var zipContents = new List<ZipFileEntry>();
            
            using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Read, leaveOpen: true))
            {
                foreach (var entry in archive.Entries)
                {
                    if (!string.IsNullOrEmpty(entry.Name))
                    {
                        var fileEntry = new ZipFileEntry
                        {
                            FileName = entry.Name,
                            FullPath = entry.FullName,
                            FileSize = entry.Length,
                            CompressedSize = entry.CompressedLength,
                            FileType = GetFileType(entry.Name),
                            LastModified = entry.LastWriteTime
                        };

                        // Opcjonalnie można odczytać zawartość pliku JSON
                        if (entry.Name.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
                        {
                            using (var reader = new StreamReader(entry.Open()))
                            {
                                fileEntry.Content = await reader.ReadToEndAsync();
                            }
                        }

                        zipContents.Add(fileEntry);
                    }
                }
            }

            var response = new ZipUploadResponse
            {
                Success = true,
                Message = "Plik ZIP został przetworzony pomyślnie",
                FileName = fileName,
                FileSize = fileSize,
                FilesCount = zipContents.Count,
                Files = zipContents,
                ProcessedAt = DateTime.UtcNow
            };

            _logger.LogTrace("Migration completed successfully!");
            
            return Ok(response);
        }
        catch (InvalidDataException ex)
        {
            _logger.LogError(ex, "Nieprawidłowy format archiwum ZIP");
            return BadRequest(new ZipUploadResponse
            {
                Success = false,
                Message = "Nieprawidłowy format archiwum ZIP",
                ErrorCode = "INVALID_ZIP_FORMAT",
                ErrorDetails = ex.Message
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Błąd podczas przetwarzania pliku ZIP");
            return StatusCode(500, new ZipUploadResponse
            {
                Success = false,
                Message = "Wystąpił błąd podczas przetwarzania pliku",
                ErrorCode = "PROCESSING_ERROR",
                ErrorDetails = ex.Message
            });
        }
    }

    private string GetFileType(string fileName)
    {
        var extension = Path.GetExtension(fileName).ToLowerInvariant();
        return extension switch
        {
            ".json" => "JSON",
            ".docx" => "Word Document",
            ".cer" or ".crt" or ".pem" or ".pfx" => "Certificate",
            _ => "Unknown"
        };
    }
}

/// <summary>
/// Zaawansowana odpowiedź z przetwarzania pliku ZIP
/// </summary>
public class ZipUploadResponse
{
    /// <summary>
    /// Czy operacja zakończyła się sukcesem
    /// </summary>
    public bool Success { get; set; }
    
    /// <summary>
    /// Komunikat dla użytkownika
    /// </summary>
    public string Message { get; set; } = string.Empty;
    
    /// <summary>
    /// Nazwa przesłanego pliku
    /// </summary>
    public string? FileName { get; set; }
    
    /// <summary>
    /// Rozmiar pliku w bajtach
    /// </summary>
    public long? FileSize { get; set; }
    
    /// <summary>
    /// Liczba plików w archiwum
    /// </summary>
    public int? FilesCount { get; set; }
    
    /// <summary>
    /// Lista plików w archiwum
    /// </summary>
    public List<ZipFileEntry>? Files { get; set; }
    
    /// <summary>
    /// Data i czas przetworzenia
    /// </summary>
    public DateTime? ProcessedAt { get; set; }
    
    /// <summary>
    /// Kod błędu (jeśli wystąpił)
    /// </summary>
    public string? ErrorCode { get; set; }
    
    /// <summary>
    /// Szczegóły błędu (jeśli wystąpił)
    /// </summary>
    public string? ErrorDetails { get; set; }
}

/// <summary>
/// Informacje o pojedynczym pliku w archiwum ZIP
/// </summary>
public class ZipFileEntry
{
    /// <summary>
    /// Nazwa pliku
    /// </summary>
    public string FileName { get; set; } = string.Empty;
    
    /// <summary>
    /// Pełna ścieżka w archiwum
    /// </summary>
    public string FullPath { get; set; } = string.Empty;
    
    /// <summary>
    /// Rozmiar rozpakowanego pliku
    /// </summary>
    public long FileSize { get; set; }
    
    /// <summary>
    /// Rozmiar skompresowanego pliku
    /// </summary>
    public long CompressedSize { get; set; }
    
    /// <summary>
    /// Typ pliku
    /// </summary>
    public string FileType { get; set; } = string.Empty;
    
    /// <summary>
    /// Data ostatniej modyfikacji
    /// </summary>
    public DateTimeOffset LastModified { get; set; }
    
    /// <summary>
    /// Zawartość pliku (opcjonalnie, dla plików JSON)
    /// </summary>
    public string? Content { get; set; }
}

