using System.Security.Claims;
using D2ViewerEditor.Api.Controllers;
using D2ViewerEditor.Api.Security;
using FluentAssertions;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Moq;
using NUnit.Framework;

namespace D2ViewerEditor.Api.UnitTests.Controllers;

/// <summary>
/// Admin Graph user lookup (empty query → 400, no result → 404, found → 200) and the resource
/// authorization endpoint (roles → allowed frontend resources; admin is a superset).
/// </summary>
[TestFixture]
public class IdentityControllerTests
{
    private static ResourcesProvider Resources() =>
        new(Options.Create(new AzureAdOptions()));

    private static IdentityController ControllerWithRoles(params string[] roles)
    {
        var identity = new ClaimsIdentity(
            roles.Select(r => new Claim("roles", r)), "test", nameType: "name", roleType: "roles");
        return new IdentityController(new DisabledGraphUserService(), Resources())
        {
            ControllerContext = new ControllerContext
            {
                HttpContext = new DefaultHttpContext { User = new ClaimsPrincipal(identity) },
            },
        };
    }

    [TestCase("")]
    [TestCase("   ")]
    public async Task Empty_query_returns_400(string query)
    {
        var controller = new IdentityController(new DisabledGraphUserService(), Resources());

        var result = await controller.FindUser(query, CancellationToken.None);

        result.Should().BeOfType<BadRequestObjectResult>();
    }

    [Test]
    public async Task No_user_found_returns_404()
    {
        // DisabledGraphUserService always returns null (Graph not configured).
        var controller = new IdentityController(new DisabledGraphUserService(), Resources());

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
        var controller = new IdentityController(graph.Object, Resources());

        var result = await controller.FindUser("john", CancellationToken.None);

        result.Should().BeOfType<OkObjectResult>().Which.Value.Should().Be(user);
    }

    [Test]
    public void Resources_admin_gets_documents_and_admin()
    {
        var result = ControllerWithRoles("Administrator").GetResources();

        var resources = result.Should().BeOfType<OkObjectResult>().Which.Value
            .Should().BeAssignableTo<IReadOnlyList<string>>().Subject;
        resources.Should().BeEquivalentTo("editor", "viewer", "admin");
    }

    [Test]
    public void Resources_operator_gets_documents_only()
    {
        var result = ControllerWithRoles("Operator").GetResources();

        var resources = result.Should().BeOfType<OkObjectResult>().Which.Value
            .Should().BeAssignableTo<IReadOnlyList<string>>().Subject;
        resources.Should().BeEquivalentTo("editor", "viewer");
        resources.Should().NotContain("admin");
    }

    [Test]
    public void Resources_user_without_app_role_gets_empty()
    {
        var result = ControllerWithRoles("SomethingElse").GetResources();

        var resources = result.Should().BeOfType<OkObjectResult>().Which.Value
            .Should().BeAssignableTo<IReadOnlyList<string>>().Subject;
        resources.Should().BeEmpty();
    }
}
