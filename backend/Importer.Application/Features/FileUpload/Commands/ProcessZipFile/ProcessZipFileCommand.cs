using Importer.Domain.Common;
using MediatR;

namespace Importer.Application.Features.FileUpload.Commands.ProcessZipFile;

/// <summary>
/// Komenda przetwarzania pliku ZIP z dokumentami
/// </summary>
public record ProcessZipFileCommand(Stream FileStream, string FileName) : IRequest<Result<ZipUploadResult>>;

public record ZipUploadResult(List<ZipFileEntryDto> Files, int TotalFiles, long TotalSize);

public record ZipFileEntryDto(string FileName, long Size, string? Content);
