using Importer.Domain.Common;
using Importer.Domain.Models;
using MediatR;

namespace Importer.Application.Features.Barcode.Commands.GenerateBarcode;

/// <summary>
/// Komenda generowania kodu kreskowego/QR
/// </summary>
public record GenerateBarcodeCommand(
    string Content,
    string BarcodeType,
    int Width,
    int Height,
    bool ShowText
) : IRequest<Result<BarcodeResponse>>;
