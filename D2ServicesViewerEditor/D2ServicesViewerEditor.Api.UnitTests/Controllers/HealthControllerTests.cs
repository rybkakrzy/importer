using System.Text.Json;
using D2ServicesViewerEditor.Api.Controllers;
using FluentAssertions;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging.Abstractions;
using NUnit.Framework;

namespace D2ServicesViewerEditor.Api.UnitTests.Controllers;

[TestFixture]
public class HealthControllerTests
{
    private static JsonElement ToJsonElement(object value)
    {
        using var document = JsonDocument.Parse(JsonSerializer.Serialize(value));
        return document.RootElement.Clone();
    }

    [Test]
    public void Get_ReturnsOk_WithExpectedPayloadShape()
    {
        var controller = new HealthController(NullLogger<HealthController>.Instance);

        var result = controller.Get();

        var ok = result.Should().BeOfType<OkObjectResult>().Subject;
        var payload = ToJsonElement(ok.Value!);

        payload.GetProperty("Status").GetString().Should().Be("Healthy");
        payload.GetProperty("Service").GetString().Should().Be("D2 Services Viewer Editor API");
        payload.GetProperty("Version").GetString().Should().Be("1.0.0");
        payload.TryGetProperty("Timestamp", out _).Should().BeTrue();
    }

    [Test]
    public void GetDetailed_ReturnsOk_WithDependenciesSection()
    {
        var controller = new HealthController(NullLogger<HealthController>.Instance);

        var result = controller.GetDetailed();

        var ok = result.Should().BeOfType<OkObjectResult>().Subject;
        var payload = ToJsonElement(ok.Value!);
        var dependencies = payload.GetProperty("Dependencies");

        payload.GetProperty("Status").GetString().Should().Be("Healthy");
        dependencies.GetProperty("Database").GetString().Should().Be("Healthy");
        dependencies.GetProperty("GoogleCloudStorage").GetString().Should().Be("Healthy");
        dependencies.GetProperty("SharedLayers").GetString().Should().Be("Healthy");
    }
}
