using D2ViewerEditor.Infrastructure.Services;
using D2ViewerEditor.Infrastructure.Services.Delivery;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Klasy opcji konfiguracyjnych — domyślne wartości i nadpisywalność (binding z konfiguracji).
/// </summary>
[TestFixture]
public class DeliveryOptionsTests
{
    [Test]
    public void DeliveryWorkerOptions_Defaults_AreSensible()
    {
        var o = new DeliveryWorkerOptions();

        o.Enabled.Should().BeTrue();
        o.PollInterval.Should().Be(TimeSpan.FromSeconds(15));
        o.BatchSize.Should().Be(20);
        o.MaxConcurrency.Should().Be(4);
        o.Lease.Should().Be(TimeSpan.FromMinutes(5));
        o.RetentionWindow.Should().Be(TimeSpan.FromHours(24));
        o.HttpTimeout.Should().Be(TimeSpan.FromSeconds(30));
        DeliveryWorkerOptions.SectionName.Should().Be("DeliveryWorker");
    }

    [Test]
    public void DeliveryWorkerOptions_CanBeOverridden()
    {
        var o = new DeliveryWorkerOptions { Enabled = false, BatchSize = 1, MaxConcurrency = 2 };

        o.Enabled.Should().BeFalse();
        o.BatchSize.Should().Be(1);
        o.MaxConcurrency.Should().Be(2);
    }

    [Test]
    public void GcsStorageOptions_Defaults_AndOverrides()
    {
        var o = new GcsStorageOptions();
        o.BucketName.Should().BeEmpty();
        o.ApiEndpoint.Should().BeNull();
        o.CredentialPath.Should().BeNull();
        GcsStorageOptions.SectionName.Should().Be("GoogleCloudStorage");

        o.BucketName = "bucket";
        o.ApiEndpoint = "http://localhost:4443";
        o.BucketName.Should().Be("bucket");
        o.ApiEndpoint.Should().Be("http://localhost:4443");
    }
}
