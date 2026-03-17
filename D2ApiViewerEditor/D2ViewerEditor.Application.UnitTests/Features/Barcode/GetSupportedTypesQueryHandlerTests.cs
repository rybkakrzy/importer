using D2ViewerEditor.Application.Features.Barcode.Queries.GetSupportedTypes;
using D2ViewerEditor.Domain.Interfaces;
using FluentAssertions;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Features.Barcode;

[TestFixture]
public class GetSupportedTypesQueryHandlerTests
{
    private IBarcodeGenerator _barcodeGenerator;
    private GetSupportedTypesQueryHandler _handler;

    [SetUp]
    public void SetUp()
    {
        _barcodeGenerator = Substitute.For<IBarcodeGenerator>();
        _handler = new GetSupportedTypesQueryHandler(_barcodeGenerator);
    }

    [Test]
    public async Task Handle_ShouldReturnSuccessResult()
    {
        // Arrange
        IReadOnlyList<string> types = new List<string> { "QRCode", "Code128" };
        _barcodeGenerator.GetSupportedTypes().Returns(types);

        // Act
        var result = await _handler.Handle(new GetSupportedTypesQuery(), CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
    }

    [Test]
    public async Task Handle_ShouldReturnTypesFromGenerator()
    {
        // Arrange
        IReadOnlyList<string> types = new List<string> { "QRCode", "Code128", "EAN13" };
        _barcodeGenerator.GetSupportedTypes().Returns(types);

        // Act
        var result = await _handler.Handle(new GetSupportedTypesQuery(), CancellationToken.None);

        // Assert
        result.Value.Should().HaveCount(3);
        result.Value.Should().Contain("QRCode");
        result.Value.Should().Contain("Code128");
        result.Value.Should().Contain("EAN13");
    }

    [Test]
    public async Task Handle_ShouldCallGetSupportedTypesOnce()
    {
        // Arrange
        IReadOnlyList<string> types = new List<string>();
        _barcodeGenerator.GetSupportedTypes().Returns(types);

        // Act
        await _handler.Handle(new GetSupportedTypesQuery(), CancellationToken.None);

        // Assert
        _barcodeGenerator.Received(1).GetSupportedTypes();
    }

    [Test]
    public async Task Handle_WhenGeneratorReturnsEmptyList_ShouldReturnEmptySuccess()
    {
        // Arrange
        IReadOnlyList<string> types = new List<string>();
        _barcodeGenerator.GetSupportedTypes().Returns(types);

        // Act
        var result = await _handler.Handle(new GetSupportedTypesQuery(), CancellationToken.None);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Value.Should().BeEmpty();
    }
}
