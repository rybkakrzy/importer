using System.IO.Compression;
using D2Tools.Domain.Common;
using D2Tools.Domain.Interfaces;
using MediatR;

namespace D2Tools.Application.Features.FileUpload.Commands.ProcessZipFile;

public class ProcessZipFileCommandHandler : IRequestHandler<ProcessZipFileCommand, Result<ZipUploadResult>>
{
    private readonly IDocxToHtmlConverter _converter;

    public ProcessZipFileCommandHandler(IDocxToHtmlConverter converter)
    {
        _converter = converter;
    }

    public async Task<Result<ZipUploadResult>> Handle(ProcessZipFileCommand request, CancellationToken cancellationToken)
    {
        using var memoryStream = new MemoryStream();
        await request.FileStream.CopyToAsync(memoryStream, cancellationToken);
        memoryStream.Position = 0;

        var files = new List<ZipFileEntryDto>();
        long totalSize = 0;

        using var archive = new ZipArchive(memoryStream, ZipArchiveMode.Read);
        foreach (var entry in archive.Entries)
        {
            if (string.IsNullOrEmpty(entry.Name)) continue;

            totalSize += entry.Length;
            string? content = null;

            if (entry.Name.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                using var entryStream = entry.Open();
                using var entryMemoryStream = new MemoryStream();
                await entryStream.CopyToAsync(entryMemoryStream, cancellationToken);
                entryMemoryStream.Position = 0;

                var documentContent = _converter.Convert(entryMemoryStream);
                content = documentContent.Html;
            }

            files.Add(new ZipFileEntryDto(entry.FullName, entry.Length, content));
        }

        var result = new ZipUploadResult(files, files.Count, totalSize);
        return Result<ZipUploadResult>.Success(result);
    }
}
