using D2ViewerEditor.Application.Common.Security;
using FluentAssertions;
using Microsoft.Extensions.Options;
using NUnit.Framework;

namespace D2ViewerEditor.Application.UnitTests.Common.Security;

[TestFixture]
public class ReturnUrlValidatorTests
{
    [Test]
    public void Validate_WithValidHttpsUrl_ReturnsNormalizedSuccess()
    {
        var validator = BuildValidator();

        var result = validator.Validate("  https://app.example.com/cb?x=1  ");

        result.IsValid.Should().BeTrue();
        result.NormalizedUrl.Should().Be("https://app.example.com/cb?x=1");
    }

    [Test]
    public void Validate_WithProtocolRelative_ReturnsFailure()
    {
        var validator = BuildValidator();

        var result = validator.Validate("//evil.example.com/cb");

        result.IsValid.Should().BeFalse();
        result.Code.Should().Be(ReturnUrlRejectionCode.ProtocolRelative);
    }

    [Test]
    public void Validate_WithEncodedScriptCharacters_ReturnsFailure()
    {
        var validator = BuildValidator();

        var result = validator.Validate("https://app.example.com/%3Cscript%3Ealert(1)%3C/script%3E");

        result.IsValid.Should().BeFalse();
        result.Code.Should().Be(ReturnUrlRejectionCode.DangerousCharacters);
    }

    [Test]
    public void Validate_WithLoopbackHost_ReturnsFailureByDefault()
    {
        var validator = BuildValidator();

        var result = validator.Validate("https://localhost/cb");

        result.IsValid.Should().BeFalse();
        result.Code.Should().Be(ReturnUrlRejectionCode.LoopbackNotAllowed);
    }

    [Test]
    public void Validate_WithPrivateIpLiteral_ReturnsFailureByDefault()
    {
        var validator = BuildValidator();

        var result = validator.Validate("https://10.1.2.3/cb");

        result.IsValid.Should().BeFalse();
        result.Code.Should().Be(ReturnUrlRejectionCode.PrivateNetworkIpNotAllowed);
    }

    [Test]
    public void Validate_WithAllowList_RejectsHostOutsideList()
    {
        var validator = BuildValidator(new ReturnUrlSecurityOptions
        {
            AllowedHosts = ["trusted.example.com"]
        });

        var result = validator.Validate("https://evil.example.com/cb");

        result.IsValid.Should().BeFalse();
        result.Code.Should().Be(ReturnUrlRejectionCode.HostNotAllowListed);
    }

    [Test]
    public void Validate_WithAllowList_AcceptsSubdomain()
    {
        var validator = BuildValidator(new ReturnUrlSecurityOptions
        {
            AllowedHosts = ["trusted.example.com"]
        });

        var result = validator.Validate("https://api.trusted.example.com/cb");

        result.IsValid.Should().BeTrue();
    }

    private static ReturnUrlValidator BuildValidator(ReturnUrlSecurityOptions? options = null)
    {
        return new ReturnUrlValidator(Options.Create(options ?? new ReturnUrlSecurityOptions()));
    }
}
