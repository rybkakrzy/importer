using D2ViewerEditor.Domain.Models;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Domain.UnitTests.Models;

[TestFixture]
public class DocumentModelsTests
{
    [Test]
    public void DocumentContent_ShouldInitializeWithDefaultValues()
    {
        // Act
        var document = new DocumentContent();

        // Assert
        document.Html.Should().BeEmpty();
        document.Metadata.Should().NotBeNull();
        document.Images.Should().NotBeNull().And.BeEmpty();
        document.Styles.Should().NotBeNull().And.BeEmpty();
        document.Header.Should().BeNull();
        document.Footer.Should().BeNull();
    }

    [Test]
    public void DocumentMetadata_ShouldAllowSettingCoreProperties()
    {
        // Arrange
        var created = DateTime.UtcNow;
        var modified = DateTime.UtcNow.AddDays(1);

        // Act
        var metadata = new DocumentMetadata
        {
            Title = "Test Document",
            Author = "John Doe",
            Subject = "Testing",
            Keywords = "test, unit, document",
            Description = "A test document",
            Created = created,
            Modified = modified,
            PageCount = 5,
            WordCount = 500
        };

        // Assert
        metadata.Title.Should().Be("Test Document");
        metadata.Author.Should().Be("John Doe");
        metadata.Subject.Should().Be("Testing");
        metadata.Keywords.Should().Be("test, unit, document");
        metadata.Description.Should().Be("A test document");
        metadata.Created.Should().Be(created);
        metadata.Modified.Should().Be(modified);
        metadata.PageCount.Should().Be(5);
        metadata.WordCount.Should().Be(500);
    }

    [Test]
    public void HeaderFooterContent_ShouldHaveDefaultValues()
    {
        // Act
        var headerFooter = new HeaderFooterContent();

        // Assert
        headerFooter.Html.Should().BeEmpty();
        headerFooter.Height.Should().Be(1.25);
        headerFooter.DifferentFirstPage.Should().BeFalse();
        headerFooter.FirstPageHtml.Should().BeNull();
    }

    [Test]
    public void DigitalSignatureInfo_ShouldAllowSettingAllProperties()
    {
        // Arrange
        var signedAt = DateTime.UtcNow;
        var validFrom = DateTime.UtcNow.AddYears(-1);
        var validTo = DateTime.UtcNow.AddYears(1);

        // Act
        var signature = new DigitalSignatureInfo
        {
            SignerName = "Jane Smith",
            SignerEmail = "jane@example.com",
            SignerTitle = "Manager",
            CertificateSubject = "CN=Jane Smith",
            CertificateIssuer = "CN=Test CA",
            CertificateSerialNumber = "123456789",
            SignedAt = signedAt,
            CertificateValidFrom = validFrom,
            CertificateValidTo = validTo,
            IsValid = true,
            ValidationMessage = "Signature is valid",
            Reason = "Approval"
        };

        // Assert
        signature.SignerName.Should().Be("Jane Smith");
        signature.SignerEmail.Should().Be("jane@example.com");
        signature.SignerTitle.Should().Be("Manager");
        signature.IsValid.Should().BeTrue();
        signature.ValidationMessage.Should().Be("Signature is valid");
        signature.Reason.Should().Be("Approval");
        signature.SignedAt.Should().Be(signedAt);
    }

    [Test]
    public void DocumentImage_ShouldAllowSettingProperties()
    {
        // Act
        var image = new DocumentImage
        {
            Id = "img1",
            ContentType = "image/png",
            Base64Data = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
        };

        // Assert
        image.Id.Should().Be("img1");
        image.ContentType.Should().Be("image/png");
        image.Base64Data.Should().NotBeEmpty();
    }
}
