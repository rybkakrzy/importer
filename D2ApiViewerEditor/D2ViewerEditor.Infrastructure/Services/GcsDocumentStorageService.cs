using D2ViewerEditor.Domain.Interfaces;
using Google.Cloud.Storage.V1;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace D2ViewerEditor.Infrastructure.Services;

/// <summary>
/// Implementacja przechowywania dokumentów w Google Cloud Storage.
/// Działa zarówno z produkcyjnym GCS jak i z fake-gcs-server (dev/test).
/// </summary>
public class GcsDocumentStorageService : IDocumentStorageService
{
    private readonly StorageClient _storageClient;
    private readonly string _bucketName;
    private readonly ILogger<GcsDocumentStorageService> _logger;

    public GcsDocumentStorageService(StorageClient storageClient, IOptions<GcsStorageOptions> options, ILogger<GcsDocumentStorageService> logger)
    {
        _storageClient = storageClient;
        _bucketName = options.Value.BucketName;
        _logger = logger;
    }

    public async Task<string> UploadAsync(Guid versionId, byte[] content, string mimeType, CancellationToken cancellationToken = default)
    {
        var objectName = BuildObjectName(versionId);

        using var stream = new MemoryStream(content);
        await _storageClient.UploadObjectAsync(
            _bucketName,
            objectName,
            mimeType,
            stream,
            cancellationToken: cancellationToken);

        _logger.LogInformation("Uploaded {ObjectName} ({Size} bytes) to bucket {Bucket}", objectName, content.Length, _bucketName);

        return objectName;
    }

    public async Task<byte[]> DownloadAsync(string storagePath, CancellationToken cancellationToken = default)
    {
        using var stream = new MemoryStream();
        await _storageClient.DownloadObjectAsync(
            _bucketName,
            storagePath,
            stream,
            cancellationToken: cancellationToken);

        return stream.ToArray();
    }

    public async Task DeleteAsync(string storagePath, CancellationToken cancellationToken = default)
    {
        try
        {
            await _storageClient.DeleteObjectAsync(_bucketName, storagePath, cancellationToken: cancellationToken);
            _logger.LogInformation("Deleted {ObjectName} from bucket {Bucket}", storagePath, _bucketName);
        }
        catch (Google.GoogleApiException ex) when (ex.HttpStatusCode == System.Net.HttpStatusCode.NotFound)
        {
            _logger.LogWarning("Object {ObjectName} not found in bucket {Bucket} during delete", storagePath, _bucketName);
        }
    }

    public async Task<bool> ExistsAsync(string storagePath, CancellationToken cancellationToken = default)
    {
        try
        {
            await _storageClient.GetObjectAsync(_bucketName, storagePath, cancellationToken: cancellationToken);
            return true;
        }
        catch (Google.GoogleApiException ex) when (ex.HttpStatusCode == System.Net.HttpStatusCode.NotFound)
        {
            return false;
        }
    }

    private static string BuildObjectName(Guid versionId)
    {
        return $"documents/{versionId}";
    }
}
