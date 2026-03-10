using D2ViewerEditor.Domain.Common;
using MediatR;

namespace D2ViewerEditor.Application.Features.Barcode.Queries.GetSupportedTypes;

/// <summary>
/// Zapytanie o listę obsługiwanych typów kodów kreskowych
/// </summary>
public record GetSupportedTypesQuery : IRequest<Result<IReadOnlyList<string>>>;
