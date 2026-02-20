using Importer.Domain.Common;
using MediatR;

namespace Importer.Application.Features.Documents.Commands.UploadImage;

/// <summary>
/// Komenda wgrywania obrazu (zwraca Base64)
/// </summary>
public record UploadImageCommand(Stream FileStream, string FileName, string ContentType, long FileSize)
    : IRequest<Result<ImageUploadResponse>>;

public record ImageUploadResponse(string Base64, string FileName, long Size);
