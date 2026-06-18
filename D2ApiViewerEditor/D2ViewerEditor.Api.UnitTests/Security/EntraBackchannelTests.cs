using System.Net;
using D2ViewerEditor.Api.Security;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Api.UnitTests.Security;

/// <summary>
/// Entra backchannel proxy builder: null when no proxy is configured (local / direct), otherwise a
/// WebProxy pointing at AzureAd:Proxy:Url with GCS bypassed.
/// </summary>
[TestFixture]
public class EntraBackchannelTests
{
    [TestCase("")]
    [TestCase("   ")]
    public void CreateProxy_returns_null_when_no_url(string url)
    {
        var options = new AzureAdOptions { Proxy = new ProxyOptions { Url = url } };

        EntraBackchannel.CreateProxy(options).Should().BeNull();
    }

    [Test]
    public void CreateProxy_returns_proxy_bypassing_gcs_when_url_set()
    {
        var options = new AzureAdOptions { Proxy = new ProxyOptions { Url = "http://corp-proxy:3128" } };

        var proxy = EntraBackchannel.CreateProxy(options);

        proxy.Should().NotBeNull();
        proxy!.Address.Should().Be(new Uri("http://corp-proxy:3128"));
        proxy.BypassList.Should().Contain("storage.googleapis.com");
        proxy.UseDefaultCredentials.Should().BeTrue();
    }

    [Test]
    public void CreateProxy_uses_explicit_credentials_when_username_set()
    {
        var options = new AzureAdOptions
        {
            Proxy = new ProxyOptions { Url = "http://corp-proxy:3128", Username = "svc", Password = "pwd" }
        };

        var proxy = EntraBackchannel.CreateProxy(options);

        proxy!.UseDefaultCredentials.Should().BeFalse();
        var credential = proxy.Credentials.Should().BeOfType<NetworkCredential>().Subject;
        credential.UserName.Should().Be("svc");
        credential.Password.Should().Be("pwd");
    }
}
