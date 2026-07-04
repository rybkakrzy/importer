using System.Net;
using System.Net.Http;
using FluentAssertions;

namespace D2ViewerEditor.Api.IntegrationTests;

/// <summary>
/// End-to-end authorization checks for the editor / document endpoints. These exercise the real
/// middleware pipeline, so they fail if the RequireAppOperator policy is ever dropped from the
/// controllers again (the fixed access-control regression). A user must hold an application role
/// (Operator or Administrator) — mere authentication is not enough.
/// </summary>
[TestFixture]
public class EditorAuthorizationTests
{
    private AuthTestWebApplicationFactory _factory = null!;

    // A representative editor endpoint whose handler is pure (static list, no DB/GCS) so a 200
    // proves the authorization pipeline let the request through — nothing else can fail.
    private const string EditorEndpoint = "/api/document/templates";

    // An admin-only endpoint (RequireAppAdmin) used to prove Operator is NOT elevated to admin.
    private const string AdminEndpoint = "/api/documentstorage";

    [OneTimeSetUp]
    public void OneTimeSetUp() => _factory = new AuthTestWebApplicationFactory();

    [OneTimeTearDown]
    public void OneTimeTearDown() => _factory.Dispose();

    private HttpClient CreateClient(string? roles)
    {
        var client = _factory.CreateClient();
        if (roles is not null)
        {
            client.DefaultRequestHeaders.Add(TestAuthHandler.AuthHeader, "true");
            if (roles.Length > 0)
                client.DefaultRequestHeaders.Add(TestAuthHandler.RolesHeader, roles);
        }
        return client;
    }

    [Test]
    public async Task EditorEndpoint_WithoutToken_Returns401()
    {
        var client = CreateClient(roles: null);

        var response = await client.GetAsync(EditorEndpoint);

        response.StatusCode.Should().Be(HttpStatusCode.Unauthorized);
    }

    [Test]
    public async Task EditorEndpoint_AuthenticatedButNoRole_Returns403()
    {
        var client = CreateClient(roles: "");

        var response = await client.GetAsync(EditorEndpoint);

        response.StatusCode.Should().Be(HttpStatusCode.Forbidden);
    }

    [Test]
    public async Task EditorEndpoint_WithUnrelatedRole_Returns403()
    {
        var client = CreateClient(roles: "SomeOtherRole");

        var response = await client.GetAsync(EditorEndpoint);

        response.StatusCode.Should().Be(HttpStatusCode.Forbidden);
    }

    [Test]
    public async Task EditorEndpoint_WithOperatorRole_Returns200()
    {
        var client = CreateClient(roles: "Operator");

        var response = await client.GetAsync(EditorEndpoint);

        response.StatusCode.Should().Be(HttpStatusCode.OK);
    }

    [Test]
    public async Task EditorEndpoint_WithAdministratorRole_Returns200()
    {
        // API-level access: RequireAppOperator admits Administrator too (the admin module reads
        // document/version data through these endpoints). Frontend AREA access (which module the user
        // may open) is governed separately by ResourcesProvider — an admin is not shown the editor UI.
        var client = CreateClient(roles: "Administrator");

        var response = await client.GetAsync(EditorEndpoint);

        response.StatusCode.Should().Be(HttpStatusCode.OK);
    }

    [Test]
    public async Task EditorSaveEndpoint_WithoutToken_Returns401()
    {
        var client = CreateClient(roles: null);

        var response = await client.PostAsync("/api/document/save",
            new StringContent("{}", System.Text.Encoding.UTF8, "application/json"));

        response.StatusCode.Should().Be(HttpStatusCode.Unauthorized);
    }

    [Test]
    public async Task EditorSaveEndpoint_AuthenticatedButNoRole_Returns403()
    {
        var client = CreateClient(roles: "");

        var response = await client.PostAsync("/api/document/save",
            new StringContent("{}", System.Text.Encoding.UTF8, "application/json"));

        response.StatusCode.Should().Be(HttpStatusCode.Forbidden);
    }

    [Test]
    public async Task AdminEndpoint_WithOperatorRole_Returns403()
    {
        // The document-storage listing requires RequireAppAdmin. Operator must NOT reach it, proving
        // the admin endpoints stay stricter than the class-level RequireAppOperator policy.
        var client = CreateClient(roles: "Operator");

        var response = await client.GetAsync(AdminEndpoint);

        response.StatusCode.Should().Be(HttpStatusCode.Forbidden);
    }

    [Test]
    public async Task AdminEndpoint_WithoutToken_Returns401()
    {
        var client = CreateClient(roles: null);

        var response = await client.GetAsync(AdminEndpoint);

        response.StatusCode.Should().Be(HttpStatusCode.Unauthorized);
    }
}
