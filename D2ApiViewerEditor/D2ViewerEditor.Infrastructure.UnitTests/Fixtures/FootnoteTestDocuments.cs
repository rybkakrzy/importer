using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace D2ViewerEditor.Infrastructure.UnitTests.Fixtures;

/// <summary>
/// Builds DOCX fixtures for footnote tests at the OOXML level, INDEPENDENTLY of the production
/// <c>HtmlToDocxConverter</c>. Testing the importer/exporter against documents produced by the
/// exporter itself would only prove the code agrees with itself; these fixtures are the known-good
/// input. Covers: plain text around a reference, ≥2 distinct footnotes, references in different
/// paragraphs, a multi-paragraph footnote, basic text formatting, Polish/Unicode text, and the
/// standard separator + continuation separator technical footnotes.
/// </summary>
internal static class FootnoteTestDocuments
{
    internal const long SeparatorId = -1;
    internal const long ContinuationSeparatorId = 0;

    /// <summary>
    /// Two distinct footnotes, references in two different paragraphs. Footnote 1 is single-paragraph
    /// with Polish/Unicode text; footnote 2 is multi-paragraph with bold + italic formatting.
    /// </summary>
    internal static byte[] TwoFootnotes()
    {
        return Build(footnotesRoot =>
        {
            AppendTechnicalSeparators(footnotesRoot);

            // Footnote 1 — single paragraph, Polish + Unicode.
            footnotesRoot.Append(UserFootnote(1,
                new Paragraph(
                    AutoNumberMarkRun(),
                    TextRun("Pierwszy przypis: zażółć gęślą jaźń — €, ©, →."))));

            // Footnote 2 — two paragraphs with bold + italic.
            footnotesRoot.Append(UserFootnote(2,
                new Paragraph(
                    AutoNumberMarkRun(),
                    TextRun("Drugi przypis, "),
                    BoldRun("pogrubiony"),
                    TextRun(" oraz "),
                    ItalicRun("kursywa"),
                    TextRun(".")),
                new Paragraph(
                    TextRun("Drugi akapit tego samego przypisu."))));
        },
        body =>
        {
            body.Append(new Paragraph(
                TextRun("Zdanie przed odwołaniem"),
                FootnoteReferenceRun(1),
                TextRun(" i tekst po odwołaniu.")));

            body.Append(new Paragraph(
                TextRun("Drugi akapit z kolejnym odwołaniem"),
                FootnoteReferenceRun(2),
                TextRun(".")));
        });
    }

    /// <summary>Regression: a document with no footnotes and no footnotes part at all.</summary>
    internal static byte[] NoFootnotes()
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(TextRun("Dokument bez przypisów.")),
                new Paragraph(TextRun("Drugi zwykły akapit."))));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>Same footnote referenced twice (from two paragraphs) — shares one content entry.</summary>
    internal static byte[] SharedFootnoteReferencedTwice()
    {
        return Build(footnotesRoot =>
        {
            AppendTechnicalSeparators(footnotesRoot);
            footnotesRoot.Append(UserFootnote(1,
                new Paragraph(AutoNumberMarkRun(), TextRun("Przypis wskazywany dwukrotnie."))));
        },
        body =>
        {
            body.Append(new Paragraph(TextRun("Pierwsze odwołanie"), FootnoteReferenceRun(1), TextRun(".")));
            body.Append(new Paragraph(TextRun("Drugie odwołanie"), FootnoteReferenceRun(1), TextRun(".")));
        });
    }

    /// <summary>
    /// A body reference to a footnote id that has no content in footnotes.xml — the importer must
    /// handle it in a controlled way (reference preserved, empty content, diagnostic) instead of
    /// crashing the whole import.
    /// </summary>
    internal static byte[] OrphanReference()
    {
        return Build(footnotesRoot =>
        {
            AppendTechnicalSeparators(footnotesRoot);
            // No user footnote for id 5.
        },
        body =>
        {
            body.Append(new Paragraph(TextRun("Odwołanie do brakującego przypisu"), FootnoteReferenceRun(5), TextRun(".")));
        });
    }

    /// <summary>
    /// One footnote plus a document-wide numbering format in settings.xml
    /// (<c>w:footnotePr/w:numFmt</c>) — used to assert the reader surfaces the format token.
    /// </summary>
    internal static byte[] FootnoteWithNumberFormat(NumberFormatValues numFmt)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;
            body.Append(new Paragraph(TextRun("Zdanie"), FootnoteReferenceRun(1), TextRun(".")));

            var footnotesPart = main.AddNewPart<FootnotesPart>();
            var footnotesRoot = new Footnotes();
            AppendTechnicalSeparators(footnotesRoot);
            footnotesRoot.Append(UserFootnote(1, new Paragraph(AutoNumberMarkRun(), TextRun("Przypis."))));
            footnotesPart.Footnotes = footnotesRoot;
            footnotesPart.Footnotes.Save();

            var settingsPart = main.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(
                new FootnoteDocumentWideProperties(new NumberingFormat { Val = numFmt }));
            settingsPart.Settings.Save();

            main.Document.Save();
        }
        return ms.ToArray();
    }

    // ----- low-level helpers -----

    private static byte[] Build(Action<Footnotes> buildFootnotes, Action<Body> buildBody)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            buildBody(body);

            var footnotesPart = main.AddNewPart<FootnotesPart>();
            var footnotesRoot = new Footnotes();
            buildFootnotes(footnotesRoot);
            footnotesPart.Footnotes = footnotesRoot;
            footnotesPart.Footnotes.Save();

            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static void AppendTechnicalSeparators(Footnotes root)
    {
        root.Append(new Footnote(new Paragraph(new Run(new SeparatorMark())))
        {
            Type = FootnoteEndnoteValues.Separator,
            Id = SeparatorId
        });
        root.Append(new Footnote(new Paragraph(new Run(new ContinuationSeparatorMark())))
        {
            Type = FootnoteEndnoteValues.ContinuationSeparator,
            Id = ContinuationSeparatorId
        });
    }

    private static Footnote UserFootnote(long id, params Paragraph[] paragraphs)
    {
        var footnote = new Footnote { Id = id };
        foreach (var paragraph in paragraphs)
            footnote.Append(paragraph);
        return footnote;
    }

    private static Run AutoNumberMarkRun() =>
        new(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
            new FootnoteReferenceMark());

    private static Run FootnoteReferenceRun(long id) =>
        new(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
            new FootnoteReference { Id = id });

    private static Run TextRun(string text) =>
        new(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

    private static Run BoldRun(string text) =>
        new(new RunProperties(new Bold()), new Text(text) { Space = SpaceProcessingModeValues.Preserve });

    private static Run ItalicRun(string text) =>
        new(new RunProperties(new Italic()), new Text(text) { Space = SpaceProcessingModeValues.Preserve });
}
