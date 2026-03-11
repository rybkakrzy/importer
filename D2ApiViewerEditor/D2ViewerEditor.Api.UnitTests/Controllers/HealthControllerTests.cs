using D2ViewerEditor.Api.Controllers;
using D2ViewerEditor.Domain.Common;
using FluentAssertions;
using MediatR;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using NSubstitute;
using NUnit.Framework;

namespace D2ViewerEditor.Api.UnitTests.Controllers;

[TestFixture]
public class HealthControllerTests
{
    private IConfiguration _configuration = null!;
    private HealthController _controller = null!;

    [SetUp]
    public void Setup()
    {
        _configuration = Substitute.For<IConfiguration>();
        _controller = new HealthController(_configuration);
    }

    [Test]
    public void GetHealth_ShouldReturnOkResult()
    {
        // Arrange
        _configuration["BuildInfo:Environment"].Returns("DEV");

        // Act
        var result = _controller.GetHealth();

        // Assert
        result.Should().BeOfType<OkObjectResult>();
    }

    [Test]
    public void GetHealth_ShouldReturnHealthResponse()
    {
        // Arrange
        _configuration["BuildInfo:Environment"].Returns("UAT");

        // Act
        var result = _controller.GetHealth() as OkObjectResult;

        // Assert
        result.Should().NotBeNull();
        var response = result!.Value as HealthResponse;
        response.Should().NotBeNull();
        response!.Status.Should().Be("healthy");
        response.Environment.Should().Be("UAT");
        response.BuildNumber.Should().NotBeNullOrEmpty();
        response.Timestamp.Should().BeCloseTo(DateTime.UtcNow, TimeSpan.FromSeconds(5));
    }

    [Test]
    public void GetHealth_WhenEnvironmentNotConfigured_ShouldReturnUnknown()
    {
        // Arrange
        _configuration["BuildInfo:Environment"].Returns((string)null!);

        // Act
        var result = _controller.GetHealth() as OkObjectResult;
        var response = result!.Value as HealthResponse;

        // Assert
        response!.Environment.Should().Be("UNKNOWN");
    }
}
