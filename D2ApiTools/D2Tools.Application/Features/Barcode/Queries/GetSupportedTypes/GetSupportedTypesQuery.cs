using D2Tools.Domain.Common;
using MediatR;

namespace D2Tools.Application.Features.Barcode.Queries.GetSupportedTypes;

/// <summary>
/// Zapytanie o listę obsługiwanych typów kodów kreskowych
/// </summary>
public record GetSupportedTypesQuery : IRequest<Result<IReadOnlyList<string>>>;
