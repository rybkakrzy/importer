using System.Security.Claims;
using D2ViewerEditor.Api.Security;
using FluentAssertions;
using Microsoft.Extensions.Options;
using NUnit.Framework;

namespace D2ViewerEditor.Api.UnitTests.Security;

/// <summary>
/// Group→role mapping (Doc2 pattern): the transformer turns Entra "groups" claims into
/// application role claims used by the authorization policies.
/// </summary>
[TestFixture]
public class Doc2ClaimsTransformerTests
{
    private static Doc2ClaimsTransformer BuildTransformer() =>
        new(Options.Create(new RolesOptions
        {
            GroupPrefix = "GSAPW4D_DOC2_",
            Roles = new List<RoleGroupMapping>
            {
                new() { RoleName = "APP_Pracownik", GroupNames = new() { "GSAPW4D_DOC2_Pracownik" } },
                new() { RoleName = "APP_Admin", GroupNames = new() { "GSAPW4D_DOC2_Administrator" } },
            }
        }));

    private static ClaimsPrincipal PrincipalWithGroups(params string[] groups)
    {
        var claims = groups.Select(g => new Claim("groups", g)).ToList();
        var identity = new ClaimsIdentity(claims, authenticationType: "TestAuth", nameType: "name", roleType: "roles");
        return new ClaimsPrincipal(identity);
    }

    [Test]
    public async Task Maps_admin_group_to_admin_role()
    {
        var principal = PrincipalWithGroups("GSAPW4D_DOC2_Administrator");

        var result = await BuildTransformer().TransformAsync(principal);

        result.IsInRole("APP_Admin").Should().BeTrue();
        result.IsInRole("APP_Pracownik").Should().BeFalse();
    }

    [Test]
    public async Task Maps_employee_group_case_insensitively()
    {
        var principal = PrincipalWithGroups("gsapw4d_doc2_pracownik");

        var result = await BuildTransformer().TransformAsync(principal);

        result.IsInRole("APP_Pracownik").Should().BeTrue();
    }

    [Test]
    public async Task No_matching_group_grants_no_role()
    {
        var principal = PrincipalWithGroups("GSAPW4D_DOC2_SomethingElse");

        var result = await BuildTransformer().TransformAsync(principal);

        result.IsInRole("APP_Admin").Should().BeFalse();
        result.IsInRole("APP_Pracownik").Should().BeFalse();
    }

    [Test]
    public async Task Is_idempotent_when_run_twice()
    {
        var transformer = BuildTransformer();
        var principal = PrincipalWithGroups("GSAPW4D_DOC2_Administrator");

        var once = await transformer.TransformAsync(principal);
        var twice = await transformer.TransformAsync(once);

        twice.FindAll("roles").Count(c => c.Value == "APP_Admin").Should().Be(1);
    }

    [Test]
    public async Task Unauthenticated_principal_is_left_untouched()
    {
        var anonymous = new ClaimsPrincipal(new ClaimsIdentity()); // not authenticated

        var result = await BuildTransformer().TransformAsync(anonymous);

        result.FindAll("roles").Should().BeEmpty();
    }

    [Test]
    public async Task Preserves_native_app_role_claims()
    {
        // App Roles and group mapping coexist: a token-issued "roles" claim survives.
        var identity = new ClaimsIdentity(
            new[] { new Claim("roles", "APP_Admin"), new Claim("groups", "GSAPW4D_DOC2_Pracownik") },
            "TestAuth", "name", "roles");
        var principal = new ClaimsPrincipal(identity);

        var result = await BuildTransformer().TransformAsync(principal);

        result.IsInRole("APP_Admin").Should().BeTrue();      // native app role preserved
        result.IsInRole("APP_Pracownik").Should().BeTrue();  // added from group mapping
    }
}
