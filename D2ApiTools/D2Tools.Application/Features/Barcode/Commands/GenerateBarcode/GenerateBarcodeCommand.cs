using D2Tools.Domain.Common;
using D2Tools.Domain.Models;
using MediatR;

namespace D2Tools.Application.Features.Barcode.Commands.GenerateBarcode;

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
