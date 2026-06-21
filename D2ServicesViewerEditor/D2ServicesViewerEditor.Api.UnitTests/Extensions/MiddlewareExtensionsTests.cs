using D2ServicesViewerEditor.Api.Extensions;
using FluentAssertions;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using NUnit.Framework;

namespace D2ServicesViewerEditor.Api.UnitTests.Extensions;

[TestFixture]
public class MiddlewareExtensionsTests
{
    [Test]
    public void UseExceptionHandlingMiddleware_ReturnsSameBuilderInstance()
    {
        var services = new ServiceCollection();
        services.AddLogging();
        var provider = services.BuildServiceProvider();
        var app = new ApplicationBuilder(provider);

        var returned = app.UseExceptionHandlingMiddleware();

        returned.Should().BeSameAs(app);
    }

    [Test]
    public void UseRequestObservability_ReturnsSameBuilderInstance()
    {
        var services = new ServiceCollection();
        services.AddLogging();
        var provider = services.BuildServiceProvider();
        var app = new ApplicationBuilder(provider);

        var returned = app.UseRequestObservability();

        returned.Should().BeSameAs(app);
    }
}
