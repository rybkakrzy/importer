using Importer.Domain.Common;
using MediatR;

namespace Importer.Application.Features.Barcode.Queries.GetSupportedTypes;

/// <summary>
/// Zapytanie o listę obsługiwanych typów kodów kreskowych
/// </summary>
public record GetSupportedTypesQuery : IRequest<Result<IReadOnlyList<string>>>;
