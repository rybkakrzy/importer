using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using D2ViewerEditor.Infrastructure.Services;
using FluentAssertions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Import DOCX z ochroną przed edycją (settings.xml). Dokument z wymuszonym
/// w:documentProtection („Ogranicz edycję") lub w:writeProtection (hasło zapisu /
/// zalecenie tylko-do-odczytu) musi wracać z konwersji z IsReadOnlyProtected=true —
/// front na tej podstawie otwiera go w trybie tylko do odczytu. Regresja zgłoszenia:
/// plik tylko-do-odczytu z Worda był w DOC2 Editor swobodnie edytowalny.
/// </summary>
[TestFixture]
public class DocumentProtectionImportTests
{
    private DocxToHtmlConverter _reader = null!;

    [SetUp]
    public void Setup()
    {
        _reader = new DocxToHtmlConverter();
    }

    private static MemoryStream Docx(OpenXmlElement? settingsChild)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Plik tylko do odczytu")))));

            if (settingsChild != null)
            {
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings(settingsChild);
                settingsPart.Settings.Save();
            }

            mainPart.Document.Save();
        }
        ms.Position = 0;
        return ms;
    }

    private static MemoryStream DocxWithMarkAsFinal(bool value)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Plik oznaczony jako ostateczny")))));
            mainPart.Document.Save();

            var customPart = doc.AddCustomFilePropertiesPart();
            customPart.Properties = new Properties(
                new CustomDocumentProperty(new VTBool(value ? "true" : "false"))
                {
                    FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                    PropertyId = 2,
                    Name = "_MarkAsFinal"
                });
            customPart.Properties.Save();
        }
        ms.Position = 0;
        return ms;
    }

    [Test]
    public void MarkAsFinal_SetsIsReadOnlyProtected()
    {
        // Word: Plik → Informacje → Chroń dokument → „Oznacz jako ostateczny" (_MarkAsFinal=true
        // w docProps/custom.xml). Word otwiera taki plik tylko do odczytu — my również.
        using var ms = DocxWithMarkAsFinal(true);

        _reader.Convert(ms).IsReadOnlyProtected.Should().BeTrue();
    }

    [Test]
    public void MarkAsFinalFalse_IsNotReadOnly()
    {
        // _MarkAsFinal=false (użytkownik cofnął oznaczenie) → dokument edytowalny.
        using var ms = DocxWithMarkAsFinal(false);

        _reader.Convert(ms).IsReadOnlyProtected.Should().BeFalse();
    }

    [Test]
    public void EnforcedReadOnlyProtection_SetsIsReadOnlyProtected()
    {
        // Word: Recenzja → Ogranicz edycję → „Bez zmian (tylko do odczytu)" + wymuszenie.
        using var ms = Docx(new DocumentProtection
        {
            Edit = DocumentProtectionValues.ReadOnly,
            Enforcement = true
        });

        _reader.Convert(ms).IsReadOnlyProtected.Should().BeTrue();
    }

    [Test]
    public void EnforcedPartialProtection_Comments_IsAlsoReadOnly()
    {
        // Trybów częściowych (komentarze/formularze) edytor nie egzekwuje → też blokada edycji.
        using var ms = Docx(new DocumentProtection
        {
            Edit = DocumentProtectionValues.Comments,
            Enforcement = true
        });

        _reader.Convert(ms).IsReadOnlyProtected.Should().BeTrue();
    }

    [Test]
    public void ProtectionWithoutEnforcement_IsNotReadOnly()
    {
        // w:enforcement=0/brak — Word NIE blokuje edycji, my też nie.
        using var ms = Docx(new DocumentProtection
        {
            Edit = DocumentProtectionValues.ReadOnly,
            Enforcement = false
        });

        _reader.Convert(ms).IsReadOnlyProtected.Should().BeFalse();
    }

    [Test]
    public void EnforcedProtectionWithEditNone_IsNotReadOnly()
    {
        using var ms = Docx(new DocumentProtection
        {
            Edit = DocumentProtectionValues.None,
            Enforcement = true
        });

        _reader.Convert(ms).IsReadOnlyProtected.Should().BeFalse();
    }

    [Test]
    public void WriteProtectionRecommended_SetsIsReadOnlyProtected()
    {
        // Word: Zapisz jako → Narzędzia → Opcje ogólne → „Zalecane tylko do odczytu".
        using var ms = Docx(new WriteProtection { Recommended = true });

        _reader.Convert(ms).IsReadOnlyProtected.Should().BeTrue();
    }

    [Test]
    public void WriteProtectionWithPasswordHash_SetsIsReadOnlyProtected()
    {
        // Hasło zapisu (nowszy wariant w:hashValue) — hasła nie weryfikujemy → tylko do odczytu.
        using var ms = Docx(new WriteProtection { HashValue = "5xK9..." });

        _reader.Convert(ms).IsReadOnlyProtected.Should().BeTrue();
    }

    [Test]
    public void EmptySettings_IsNotReadOnly()
    {
        using var ms = Docx(new EvenAndOddHeaders());

        _reader.Convert(ms).IsReadOnlyProtected.Should().BeFalse();
    }

    [Test]
    public void NoSettingsPart_IsNotReadOnly()
    {
        using var ms = Docx(null);

        _reader.Convert(ms).IsReadOnlyProtected.Should().BeFalse();
    }
}
