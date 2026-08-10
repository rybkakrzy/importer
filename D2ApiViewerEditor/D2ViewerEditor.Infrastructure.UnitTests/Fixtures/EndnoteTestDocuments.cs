using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace D2ViewerEditor.Infrastructure.UnitTests.Fixtures;

/// <summary>
/// Builds DOCX fixtures for endnote tests at the OOXML level, INDEPENDENTLY of the production
/// <c>HtmlToDocxConverter</c> (same rationale as <see cref="FootnoteTestDocuments"/>: the input
/// must be known-good, not produced by the code under test). Covers: a single endnote, multiple
/// distinct endnotes across paragraphs, a multi-paragraph endnote with formatting, a shared
/// endnote referenced twice, an orphan reference, a document with BOTH footnotes and endnotes
/// (distinct semantics/parts), and the standard separator + continuation separator technical
/// endnotes.
/// </summary>
internal static class EndnoteTestDocuments
{
    internal const long SeparatorId = -1;
    internal const long ContinuationSeparatorId = 0;

    /// <summary>Single endnote, single reference — the minimal case.</summary>
    internal static byte[] SingleEndnote()
    {
        return Build(endnotesRoot =>
        {
            AppendTechnicalSeparators(endnotesRoot);
            endnotesRoot.Append(UserEndnote(1,
                new Paragraph(AutoNumberMarkRun(), TextRun("Jedyny przypis końcowy."))));
        },
        body =>
        {
            body.Append(new Paragraph(
                TextRun("Zdanie przed odwołaniem"),
                EndnoteReferenceRun(1),
                TextRun(" i dalej.")));
        });
    }

    /// <summary>
    /// Two distinct endnotes, references in two different paragraphs. Endnote 1 is single-paragraph
    /// with Polish/Unicode text; endnote 2 is multi-paragraph with bold + italic formatting.
    /// </summary>
    internal static byte[] TwoEndnotes()
    {
        return Build(endnotesRoot =>
        {
            AppendTechnicalSeparators(endnotesRoot);

            endnotesRoot.Append(UserEndnote(1,
                new Paragraph(
                    AutoNumberMarkRun(),
                    TextRun("Pierwszy przypis końcowy: zażółć gęślą jaźń — €, ©, →."))));

            endnotesRoot.Append(UserEndnote(2,
                new Paragraph(
                    AutoNumberMarkRun(),
                    TextRun("Drugi przypis końcowy, "),
                    BoldRun("pogrubiony"),
                    TextRun(" oraz "),
                    ItalicRun("kursywa"),
                    TextRun(".")),
                new Paragraph(
                    TextRun("Drugi akapit tego samego przypisu końcowego."))));
        },
        body =>
        {
            body.Append(new Paragraph(
                TextRun("Zdanie przed odwołaniem"),
                EndnoteReferenceRun(1),
                TextRun(" i tekst po odwołaniu.")));

            body.Append(new Paragraph(
                TextRun("Drugi akapit z kolejnym odwołaniem"),
                EndnoteReferenceRun(2),
                TextRun(".")));
        });
    }

    /// <summary>Same endnote referenced twice (from two paragraphs) — shares one content entry.</summary>
    internal static byte[] SharedEndnoteReferencedTwice()
    {
        return Build(endnotesRoot =>
        {
            AppendTechnicalSeparators(endnotesRoot);
            endnotesRoot.Append(UserEndnote(1,
                new Paragraph(AutoNumberMarkRun(), TextRun("Przypis końcowy wskazywany dwukrotnie."))));
        },
        body =>
        {
            body.Append(new Paragraph(TextRun("Pierwsze odwołanie"), EndnoteReferenceRun(1), TextRun(".")));
            body.Append(new Paragraph(TextRun("Drugie odwołanie"), EndnoteReferenceRun(1), TextRun(".")));
        });
    }

    /// <summary>
    /// A body reference to an endnote id that has no content in endnotes.xml — the importer must
    /// preserve the reference with empty content and a diagnostic, not crash.
    /// </summary>
    internal static byte[] OrphanReference()
    {
        return Build(endnotesRoot =>
        {
            AppendTechnicalSeparators(endnotesRoot);
            // No user endnote for id 5.
        },
        body =>
        {
            body.Append(new Paragraph(TextRun("Odwołanie do brakującego przypisu końcowego"), EndnoteReferenceRun(5), TextRun(".")));
        });
    }

    /// <summary>Regression: a document with no endnotes and no endnotes part at all.</summary>
    internal static byte[] NoEndnotes()
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(TextRun("Dokument bez przypisów końcowych.")),
                new Paragraph(TextRun("Drugi zwykły akapit."))));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>
    /// A document with BOTH a footnote and an endnote referenced from the body. Proves the two
    /// types keep distinct references (w:footnoteReference vs w:endnoteReference), distinct parts
    /// and independent numbering — the reader must not collapse or swap them.
    /// </summary>
    internal static byte[] FootnoteAndEndnote()
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            body.Append(new Paragraph(
                TextRun("Tekst z przypisem dolnym"),
                FootnoteReferenceRun(1),
                TextRun(" oraz końcowym"),
                EndnoteReferenceRun(1),
                TextRun(".")));

            var footnotesPart = main.AddNewPart<FootnotesPart>();
            var footnotesRoot = new Footnotes();
            footnotesRoot.Append(SeparatorFootnote());
            footnotesRoot.Append(ContinuationSeparatorFootnote());
            footnotesRoot.Append(new Footnote(new Paragraph(AutoNumberMarkRun(), TextRun("Treść przypisu DOLNEGO."))) { Id = 1 });
            footnotesPart.Footnotes = footnotesRoot;
            footnotesPart.Footnotes.Save();

            var endnotesPart = main.AddNewPart<EndnotesPart>();
            var endnotesRoot = new Endnotes();
            AppendTechnicalSeparators(endnotesRoot);
            endnotesRoot.Append(UserEndnote(1, new Paragraph(AutoNumberMarkRun(), TextRun("Treść przypisu KOŃCOWEGO."))));
            endnotesPart.Endnotes = endnotesRoot;
            endnotesPart.Endnotes.Save();

            main.Document.Save();
        }
        return ms.ToArray();
    }

    // ----- low-level helpers -----

    /// <summary>
    /// One endnote plus a document-wide numbering format in settings.xml
    /// (<c>w:endnotePr/w:numFmt</c>) — used to assert the reader surfaces the format token.
    /// </summary>
    internal static byte[] EndnoteWithNumberFormat(NumberFormatValues numFmt)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;
            body.Append(new Paragraph(TextRun("Zdanie"), EndnoteReferenceRun(1), TextRun(".")));

            var endnotesPart = main.AddNewPart<EndnotesPart>();
            var endnotesRoot = new Endnotes();
            AppendTechnicalSeparators(endnotesRoot);
            endnotesRoot.Append(UserEndnote(1, new Paragraph(AutoNumberMarkRun(), TextRun("Przypis końcowy."))));
            endnotesPart.Endnotes = endnotesRoot;
            endnotesPart.Endnotes.Save();

            var settingsPart = main.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(
                new EndnoteDocumentWideProperties(new NumberingFormat { Val = numFmt }));
            settingsPart.Settings.Save();

            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>Format numeracji w SEKCYJNYM w:sectPr/w:endnotePr (bez settings.xml) —
    /// Word daje takiemu override'owi pierwszeństwo nad document-wide.</summary>
    internal static byte[] EndnoteWithSectionNumberFormat(NumberFormatValues numFmt)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;
            body.Append(new Paragraph(TextRun("Zdanie"), EndnoteReferenceRun(1), TextRun(".")));
            body.Append(new SectionProperties(
                new EndnoteProperties(new NumberingFormat { Val = numFmt }),
                new PageSize { Width = 11906, Height = 16838 }));

            var endnotesPart = main.AddNewPart<EndnotesPart>();
            var endnotesRoot = new Endnotes();
            AppendTechnicalSeparators(endnotesRoot);
            endnotesRoot.Append(UserEndnote(1, new Paragraph(AutoNumberMarkRun(), TextRun("Przypis końcowy."))));
            endnotesPart.Endnotes = endnotesRoot;
            endnotesPart.Endnotes.Save();

            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static byte[] Build(Action<Endnotes> buildEndnotes, Action<Body> buildBody)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            buildBody(body);

            var endnotesPart = main.AddNewPart<EndnotesPart>();
            var endnotesRoot = new Endnotes();
            buildEndnotes(endnotesRoot);
            endnotesPart.Endnotes = endnotesRoot;
            endnotesPart.Endnotes.Save();

            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static void AppendTechnicalSeparators(Endnotes root)
    {
        root.Append(new Endnote(new Paragraph(new Run(new SeparatorMark())))
        {
            Type = FootnoteEndnoteValues.Separator,
            Id = SeparatorId
        });
        root.Append(new Endnote(new Paragraph(new Run(new ContinuationSeparatorMark())))
        {
            Type = FootnoteEndnoteValues.ContinuationSeparator,
            Id = ContinuationSeparatorId
        });
    }

    private static Footnote SeparatorFootnote() =>
        new(new Paragraph(new Run(new SeparatorMark()))) { Type = FootnoteEndnoteValues.Separator, Id = SeparatorId };

    private static Footnote ContinuationSeparatorFootnote() =>
        new(new Paragraph(new Run(new ContinuationSeparatorMark()))) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = ContinuationSeparatorId };

    private static Endnote UserEndnote(long id, params Paragraph[] paragraphs)
    {
        var endnote = new Endnote { Id = id };
        foreach (var paragraph in paragraphs)
            endnote.Append(paragraph);
        return endnote;
    }

    private static Run AutoNumberMarkRun() =>
        new(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
            new EndnoteReferenceMark());

    private static Run EndnoteReferenceRun(long id) =>
        new(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
            new EndnoteReference { Id = id });

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
