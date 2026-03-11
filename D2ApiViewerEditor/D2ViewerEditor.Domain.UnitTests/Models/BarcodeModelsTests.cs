using D2ViewerEditor.Domain.Models;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Domain.UnitTests.Models;

[TestFixture]
public class BarcodeModelsTests
{
    [Test]
    public void BarcodeRequest_ShouldHaveDefaultValues()
    {
        // Act
        var request = new BarcodeRequest();

        // Assert
        request.Content.Should().BeEmpty();
        request.BarcodeType.Should().Be("QRCode");
        request.Width.Should().Be(300);
        request.Height.Should().Be(300);
        request.ShowText.Should().BeFalse();
    }

    [Test]
    public void BarcodeRequest_ShouldAllowSettingProperties()
    {
        // Arrange & Act
        var request = new BarcodeRequest
        {
            Content = "TEST123",
            BarcodeType = "Code128",
            Width = 500,
            Height = 200,
            ShowText = true
        };

        // Assert
        request.Content.Should().Be("TEST123");
        request.BarcodeType.Should().Be("Code128");
        request.Width.Should().Be(500);
        request.Height.Should().Be(200);
        request.ShowText.Should().BeTrue();
    }

    [Test]
    public void BarcodeResponse_ShouldHaveDefaultValues()
    {
        // Act
        var response = new BarcodeResponse();

        // Assert
        response.Base64Image.Should().BeEmpty();
        response.ContentType.Should().Be("image/png");
        response.BarcodeType.Should().BeEmpty();
    }

    [Test]
    public void BarcodeResponse_ShouldAllowSettingProperties()
    {
        // Arrange & Act
        var response = new BarcodeResponse
        {
            Base64Image = "data:image/png;base64,iVBORw0KGgo...",
            ContentType = "image/png",
            BarcodeType = "QRCode"
        };

        // Assert
        response.Base64Image.Should().Be("data:image/png;base64,iVBORw0KGgo...");
        response.ContentType.Should().Be("image/png");
        response.BarcodeType.Should().Be("QRCode");
    }
}
