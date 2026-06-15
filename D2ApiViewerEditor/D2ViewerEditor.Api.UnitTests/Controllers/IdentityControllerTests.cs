using D2ViewerEditor.Api.Controllers;
using D2ViewerEditor.Api.Security;
using FluentAssertions;
using Microsoft.AspNetCore.Mvc;
using Moq;
using NUnit.Framework;

namespace D2ViewerEditor.Api.UnitTests.Controllers;

/// <summary>
/// Admin Graph user lookup endpoint. Empty query → 400, no result (incl. Graph disabled) → 404,
/// found → 200 with payload.
/// </summary>
[TestFixture]
public class IdentityControllerTests
{
    [TestCase("")]
    [TestCase("   ")]
    public async Task Empty_query_returns_400(string query)
    {
        var controller = new IdentityController(new DisabledGraphUserService());

        var result = await controller.FindUser(query, CancellationToken.None);

        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task No_user_found_returns_404()
    {
        // DisabledGraphUserService always returns null (Graph not configured).
        var controller = new IdentityController(new DisabledGraphUserService());

        var result = await controller.FindUser("john", CancellationToken.None);

        result.Should().BeOfType<NotFoundResult>();
    }

    [Test]
    public async Task Found_user_returns_200_with_payload()
    {
        var user = new GraphUserInfo("id-1", "jd@example.com", "jd@example.com",
            "John Doe", "John", "Doe", "Developer", "IT");
        var graph = new Mock<IGraphUserService>();
        graph.Setup(g => g.FindUserAsync("john", It.IsAny<CancellationToken>())).ReturnsAsync(user);
        var controller = new IdentityController(graph.Object);

        var result = await controller.FindUser("john", CancellationToken.None);

        result.Should().BeOfType<OkObjectResult>().Which.Value.Should().Be(user);
    }
}
