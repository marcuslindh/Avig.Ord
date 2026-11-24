using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace Avig.Ord.Tests;

public class ReaderTests
{
    private const string Styles = """
        <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
        <w:style w:type="paragraph" w:styleId="Rubrik1"><w:name w:val="Rubrik 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Underrubrik"><w:name w:val="Min underrubrik"/><w:basedOn w:val="Rubrik1"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/></w:style>
        <w:style w:type="character" w:styleId="Rod"><w:name w:val="Röd"/><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>
        <w:style w:type="paragraph" w:styleId="Loop1"><w:name w:val="Loop 1"/><w:basedOn w:val="Loop2"/></w:style>
        <w:style w:type="paragraph" w:styleId="Loop2"><w:name w:val="Loop 2"/><w:basedOn w:val="Loop1"/></w:style>
        """;

    [Fact]
    public void The_validator_finds_an_element_out_of_place()
    {
        byte[] file = TestFiles.Docx("<w:p><w:r><w:t>x</w:t><w:rPr><w:b/></w:rPr></w:r></w:p>");

        Assert.NotEqual(string.Empty, TestFiles.Validate(file));
    }

    [Fact]
    public async Task Formatting_is_inherited_through_the_style_chain()
    {
        byte[] file = TestFiles.Docx(
            """
            <w:p><w:pPr><w:pStyle w:val="Underrubrik"/></w:pPr><w:r><w:t>Ärv</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Rod"/><w:b w:val="0"/></w:rPr><w:t>röd</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="Rubrik1"/><w:outlineLvl w:val="9"/></w:pPr><w:r><w:t>Inte rubrik</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Från namnet</w:t></w:r></w:p>
            <w:p><w:pPr><w:pStyle w:val="Loop1"/></w:pPr><w:r><w:t>Loop</w:t></w:r></w:p>
            """,
            Styles);
        OrdDocument document = await TestFiles.OpenAsync(file);
        IReadOnlyList<Paragraph> paragraphs = document.Paragraphs;

        Run inherited = paragraphs[0].Runs[0];
        Assert.True(inherited.Format.Bold);
        Assert.True(inherited.Format.Italic);
        Assert.Equal(Length.Pt(16), inherited.Format.Size);
        Assert.Null(inherited.DirectFormat.Bold);
        Assert.Equal(1, paragraphs[0].OutlineLevel);
        Assert.Equal("Min underrubrik", paragraphs[0].StyleName);

        Run red = paragraphs[0].Runs[1];
        Assert.Equal("FF0000", red.Format.Color);
        Assert.False(red.Format.Bold);

        Assert.Null(paragraphs[1].OutlineLevel);
        Assert.Equal(2, paragraphs[2].OutlineLevel);
        Assert.Null(paragraphs[3].OutlineLevel);
        Assert.True(document.Styles.TryFindByName("rubrik 1", out Style? style));
        Assert.Equal("Rubrik1", style.Id);
    }

    [Fact]
    public async Task List_labels_follow_the_definitions_and_go_on_across_instances()
    {
        const string numbering = """
            <w:abstractNum w:abstractNumId="7">
              <w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1)"/></w:lvl>
              <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2"/></w:lvl>
            </w:abstractNum>
            <w:abstractNum w:abstractNumId="8"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/></w:lvl></w:abstractNum>
            <w:num w:numId="1"><w:abstractNumId w:val="7"/></w:num>
            <w:num w:numId="2"><w:abstractNumId w:val="7"/></w:num>
            <w:num w:numId="3"><w:abstractNumId w:val="7"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>
            <w:num w:numId="4"><w:abstractNumId w:val="8"/></w:num>
            """;
        string item(string id, int level, string text) => $"<w:p><w:pPr><w:numPr><w:ilvl w:val=\"{level}\"/><w:numId w:val=\"{id}\"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";

        byte[] file = TestFiles.Docx(item("1", 0, "a") + item("1", 1, "b") + item("1", 1, "c") + item("2", 0, "d") + item("2", 1, "e") + item("3", 0, "f") + item("4", 0, "g") + item("0", 0, "h"), null, numbering);
        OrdDocument document = await TestFiles.OpenAsync(file);

        Assert.Equal(["III)", "III.a", "III.b", "IV)", "IV.a", "I)", "•", "-"], document.Paragraphs.Select(paragraph => paragraph.List?.Label ?? "-"));
        Assert.Equal("III) a\nIII.a b\nIII.b c\nIV) d\nIV.a e\nI) f\n• g\nh", document.ToPlainText());
    }

    [Theory]
    [InlineData(RevisionView.Accepted, "Det var en ny text.")]
    [InlineData(RevisionView.Original, "Det var en gammal text.")]
    [InlineData(RevisionView.Marked, "Det var en gammalny text.")]
    public async Task Tracked_changes_are_read_in_the_view_asked_for(RevisionView view, string expected)
    {
        byte[] file = TestFiles.Docx(TrackedBody);

        OrdDocument document = await TestFiles.OpenAsync(file, new ReadOptions { Revisions = view });

        Assert.Equal(expected, document.Paragraphs[0].Text);

        if (view == RevisionView.Marked)
        {
            Assert.Equal(RevisionKind.Deleted, document.Paragraphs[0].Runs.Single(run => run.Text == "gammal").Revision);
            Assert.Equal(RevisionKind.Inserted, document.Paragraphs[0].Runs.Single(run => run.Text == "ny").Revision);
        }
    }

    [Fact]
    public async Task Accepting_and_rejecting_changes_leave_plain_text()
    {
        OrdDocument accepted = await TestFiles.OpenAsync(TestFiles.Docx(TrackedBody), new ReadOptions { Revisions = RevisionView.Marked });
        accepted.AcceptAllRevisions();
        OrdDocument rejected = await TestFiles.OpenAsync(TestFiles.Docx(TrackedBody), new ReadOptions { Revisions = RevisionView.Marked });
        rejected.RejectAllRevisions();

        Assert.Equal("Det var en ny text.", accepted.ToPlainText());
        Assert.Equal("Det var en gammal text.", rejected.ToPlainText());

        foreach (OrdDocument document in (OrdDocument[])[accepted, rejected])
        {
            byte[] saved = await TestFiles.SaveAsync(document);
            string xml = Encoding.UTF8.GetString(await PartAsync(saved, "word/document.xml"));
            Assert.DoesNotContain("w:ins", xml, StringComparison.Ordinal);
            Assert.DoesNotContain("w:del", xml, StringComparison.Ordinal);
            Assert.Equal(string.Empty, TestFiles.Validate(saved));
        }
    }

    [Fact]
    public async Task Fields_give_their_result_and_keep_their_instruction()
    {
        byte[] file = TestFiles.Docx(
            """
            <w:p>
              <w:r><w:t xml:space="preserve">Kära </w:t></w:r>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> MERGEFIELD "Förnamn" \* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>«Förnamn»</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
              <w:fldSimple w:instr=" PAGE "><w:r><w:t>4</w:t></w:r></w:fldSimple>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>HYPERLINK "https://example.com/"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>länk</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>IF </w:instrText></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>MERGEFIELD X</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>inre</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:instrText> = 1 "ja" "nej"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>ja</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
            </w:p>
            """);
        OrdDocument document = await TestFiles.OpenAsync(file);
        Paragraph paragraph = document.Paragraphs[0];
        List<Field> fields = [.. paragraph.Inlines.OfType<Field>()];

        Assert.Equal("Kära «Förnamn»4länkja", paragraph.Text);
        Assert.Equal(4, fields.Count);
        Assert.Equal("Förnamn", fields[0].MergeFieldName);
        Assert.Equal("PAGE", fields[1].Code);
        Assert.Equal("ja", fields[3].Text);
        Assert.Contains("[länk](https://example.com/)", document.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Content_controls_text_boxes_and_alternate_content_are_read()
    {
        byte[] file = TestFiles.Docx(
            """
            <w:sdt><w:sdtPr><w:tag w:val="block"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>I kontroll</w:t></w:r></w:p></w:sdtContent></w:sdt>
            <w:p>
              <w:sdt><w:sdtPr><w:alias w:val="Namn"/><w:tag w:val="namn"/></w:sdtPr><w:sdtContent><w:r><w:t>Anna</w:t></w:r></w:sdtContent></w:sdt>
              <w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wps:txbx><w:txbxContent><w:p><w:r><w:t>I rutan</w:t></w:r></w:p></w:txbxContent></wps:txbx></w:drawing></mc:Choice><mc:Fallback><w:pict><w:p><w:r><w:t>Reserv</w:t></w:r></w:p></w:pict></mc:Fallback></mc:AlternateContent></w:r>
            </w:p>
            """);
        OrdDocument document = await TestFiles.OpenAsync(file);

        Assert.Equal("I kontroll", document.Paragraphs[0].Text);
        ContentControl control = Assert.Single(document.Paragraphs[1].Inlines.OfType<ContentControl>());
        Assert.Equal("namn", control.Tag);
        Assert.Equal("Namn", control.Title);
        Assert.Equal("I kontroll\nAnna\nI rutan", document.ToPlainText());
    }

    [Fact]
    public async Task A_strict_document_is_read_and_saved_as_strict()
    {
        byte[] file = TestFiles.Docx("<w:p><w:r><w:t>Strikt</w:t></w:r></w:p>");
        byte[] strict = await ReplaceInPartAsync(file, "word/document.xml", TestFiles.Wml, "http://purl.oclc.org/ooxml/wordprocessingml/main");

        OrdDocument document = await TestFiles.OpenAsync(strict);
        Assert.Equal("Strikt", document.ToPlainText());
        document.AddParagraph("Mer");
        byte[] saved = await TestFiles.SaveAsync(document);

        Assert.Contains("purl.oclc.org/ooxml/wordprocessingml/main", Encoding.UTF8.GetString(await PartAsync(saved, "word/document.xml")), StringComparison.Ordinal);
        Assert.Equal("Strikt\nMer", (await TestFiles.OpenAsync(saved)).ToPlainText());
    }

    [Fact]
    public async Task The_streaming_reader_gives_the_same_text_as_the_document_model()
    {
        byte[] file = await WriterTests.EveryFeatureAsync(false);
        OrdDocument document = await TestFiles.OpenAsync(file);
        Result<DocxReader> opened = await DocxReader.OpenAsync(new MemoryStream(file));
        Assert.True(opened.IsSuccess(out DocxReader? reader));

        await using (reader)
        {
            List<Block> blocks = [];

            await foreach (Block block in reader.ReadBlocksAsync())
            {
                Assert.True(block.IsReadOnly);
                blocks.Add(block);
            }

            Assert.Equal(document.Blocks.Count, blocks.Count);
            Assert.Equal(document.Blocks.Select(block => block.Text), blocks.Select(block => block.Text));
            Assert.Equal(document.Paragraphs.Select(paragraph => paragraph.List?.Label), OrdDocument.AllParagraphs(blocks).Select(paragraph => paragraph.List?.Label));
            Assert.Throws<InvalidOperationException>(() => ((Paragraph)blocks[0]).AppendText("x"));
        }
    }

    [Fact]
    public async Task The_event_reader_gives_the_same_text_as_the_paragraphs()
    {
        byte[] file = await WriterTests.EveryFeatureAsync(false);
        OrdDocument document = await TestFiles.OpenAsync(file);
        Result<DocxReader> opened = await DocxReader.OpenAsync(new MemoryStream(file));
        Assert.True(opened.IsSuccess(out DocxReader? reader));
        List<string> paragraphs = [];
        StringBuilder current = new();
        int tables = 0;

        await using (reader)
        {
            await foreach (DocumentEvent item in reader.ReadEventsAsync())
            {
                switch (item.Kind)
                {
                    case DocumentEventKind.ParagraphStart:
                        current.Clear();

                        break;

                    case DocumentEventKind.Text:
                        current.Append(item.Text);

                        break;

                    case DocumentEventKind.Break:
                        current.Append('\n');

                        break;

                    case DocumentEventKind.ParagraphEnd:
                        paragraphs.Add(current.ToString());

                        break;

                    case DocumentEventKind.TableStart:
                        tables++;

                        break;

                    default:
                        break;
                }
            }
        }

        Assert.Equal(1, tables);
        Assert.Equal(document.Paragraphs.Select(paragraph => paragraph.Text), paragraphs);
    }

    [Fact]
    public async Task Markdown_has_headings_lists_tables_emphasis_links_and_pictures()
    {
        byte[] file = await WriterTests.EveryFeatureAsync(false);
        OrdDocument document = await TestFiles.OpenAsync(file);

        string markdown = document.ToMarkdown();

        Assert.Contains("# Årsrapport\n", markdown, StringComparison.Ordinal);
        Assert.Contains("Vanlig **fet** *kursiv* efter tabb  \nny rad", markdown, StringComparison.Ordinal);
        Assert.Contains("[Läs mer](https://example.com/)", markdown, StringComparison.Ordinal);
        Assert.Contains("- Ett\n- Två\n    - Under\n1. Först\n2. Sedan\n1. Ny lista\n", markdown, StringComparison.Ordinal);
        Assert.Contains("| **Namn** | **Belopp** | **Not** |\n| --- | --- | --- |\n| Hyra | 8 000 | Hela |\n| Summa |  |  |", markdown, StringComparison.Ordinal);
        Assert.Contains("![En röd punkt](media/image1.png)", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Markdown_escapes_what_would_be_read_as_markdown()
    {
        byte[] file = TestFiles.Docx("<w:p><w:r><w:t>1. inte en lista *och* [inte] en länk_</w:t></w:r></w:p><w:p><w:r><w:t># inte rubrik</w:t></w:r></w:p>");

        string markdown = (await TestFiles.OpenAsync(file)).ToMarkdown();

        Assert.Equal("1\\. inte en lista \\*och\\* \\[inte\\] en länk\\_\n\n\\# inte rubrik\n", markdown);
    }

    [Fact]
    public async Task Parts_the_package_does_not_know_are_saved_unchanged()
    {
        byte[] file = TestFiles.Docx(
            "<w:p><w:r><w:t>Text</w:t></w:r><w:unknownElement xmlns:x=\"urn:x\"/></w:p>",
            extra: main =>
            {
                CustomXmlPart custom = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                using StreamWriter writer = new(custom.GetStream(FileMode.Create));
                writer.Write("<data xmlns=\"urn:test\"><value>42</value></data>");
            });
        OrdDocument document = await TestFiles.OpenAsync(file);
        document.AddParagraph("Ny");

        byte[] saved = await TestFiles.SaveAsync(document);

        using WordprocessingDocument reopened = WordprocessingDocument.Open(new MemoryStream(saved), false);
        CustomXmlPart part = Assert.Single(reopened.MainDocumentPart!.CustomXmlParts);
        using StreamReader reader = new(part.GetStream());
        Assert.Contains("<value>42</value>", await reader.ReadToEndAsync(), StringComparison.Ordinal);
        Assert.Contains("unknownElement", Encoding.UTF8.GetString(await PartAsync(saved, "word/document.xml")), StringComparison.Ordinal);
        Assert.Equal("Text\nNy", TestFiles.SdkText(saved));
    }

    [Fact]
    public async Task A_document_opened_and_saved_unchanged_keeps_its_bytes_for_parts_never_read()
    {
        byte[] file = await WriterTests.EveryFeatureAsync(false);
        OrdDocument document = await TestFiles.OpenAsync(file);

        byte[] saved = await TestFiles.SaveAsync(document);

        Assert.Equal(await PartAsync(file, "word/media/image1.png"), await PartAsync(saved, "word/media/image1.png"));
        Assert.Equal(await PartAsync(file, "word/settings.xml"), await PartAsync(saved, "word/settings.xml"));
        Assert.Equal(string.Empty, TestFiles.Validate(saved));
    }

    [Fact]
    public async Task Macros_are_kept_or_removed_but_never_run()
    {
        byte[] file = TestFiles.Docx("<w:p><w:r><w:t>Makro</w:t></w:r></w:p>");
        byte[] macro = await ConvertToMacroAsync(file);
        OrdDocument document = await TestFiles.OpenAsync(macro);

        Assert.True(document.HasMacros);
        Assert.Equal(WordFileKind.MacroEnabledDocument, document.Kind);

        OrdDocument kept = await TestFiles.OpenAsync(await TestFiles.SaveAsync(document));
        Assert.True(kept.HasMacros);

        document.RemoveMacros();
        OrdDocument removed = await TestFiles.OpenAsync(await TestFiles.SaveAsync(document));
        Assert.False(removed.HasMacros);
        Assert.Equal(WordFileKind.Document, removed.Kind);
    }

    [Fact]
    public async Task A_document_can_be_saved_as_a_template_and_back()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:t>Mall</w:t></w:r></w:p>"));

        OrdDocument template = await TestFiles.OpenAsync(await TestFiles.SaveAsync(document, new SaveOptions { Format = WordSaveFormat.Dotx }));
        OrdDocument again = await TestFiles.OpenAsync(await TestFiles.SaveAsync(template, new SaveOptions { Format = WordSaveFormat.Docx }));

        Assert.Equal(WordFileKind.Template, template.Kind);
        Assert.Equal(WordFileKind.Document, again.Kind);
        Assert.Equal("Mall", again.ToPlainText());
    }

    [Fact]
    public async Task A_spreadsheet_is_not_a_word_document()
    {
        using MemoryStream file = new();

        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(file, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            WorkbookPart workbook = spreadsheet.AddWorkbookPart();
            workbook.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
        }

        Result<OrdDocument> opened = await OrdDocument.OpenAsync(new MemoryStream(file.ToArray()));

        Assert.Equal(OrdErrorKind.NotWord, opened.ErrorKind());
    }

    [Fact]
    public async Task Errors_have_kinds()
    {
        Assert.Equal(OrdErrorKind.NotFound, (await OrdDocument.OpenAsync(Path.Combine(Path.GetTempPath(), "does-not-exist-" + Guid.NewGuid() + ".docx"))).ErrorKind());
        Assert.Equal(OrdErrorKind.NotWord, (await OrdDocument.OpenAsync(new MemoryStream([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0]))).ErrorKind());
        Assert.Equal(OrdErrorKind.NotWord, (await OrdDocument.OpenAsync(new MemoryStream("plain text"u8.ToArray()))).ErrorKind());
        Assert.Equal(OrdErrorKind.TooLarge, (await OrdDocument.OpenAsync(new MemoryStream(TestFiles.Docx("<w:p/>")), new ReadOptions { Limits = OrdLimits.Default with { MaxUncompressedBytes = 100 } })).ErrorKind());
    }

    [Fact]
    public async Task Headers_footers_notes_comments_and_properties_are_read()
    {
        byte[] file = TestFiles.Docx(
            "<w:p><w:commentRangeStart w:id=\"0\"/><w:r><w:t>Kommenterad</w:t></w:r><w:commentRangeEnd w:id=\"0\"/><w:r><w:commentReference w:id=\"0\"/></w:r></w:p>",
            extra: main =>
            {
                WordprocessingCommentsPart comments = main.AddNewPart<WordprocessingCommentsPart>();
                TestFiles.Write(comments, $"<w:comments xmlns:w=\"{TestFiles.Wml}\"><w:comment w:id=\"0\" w:author=\"Eva\" w:initials=\"E\" w:date=\"2026-11-24T10:00:00Z\"><w:p><w:r><w:t>Stämmer detta?</w:t></w:r></w:p></w:comment></w:comments>");
            });
        OrdDocument document = await TestFiles.OpenAsync(file);

        Comment comment = Assert.Single(document.Comments);
        Assert.Equal("Eva", comment.Author);
        Assert.Equal("Stämmer detta?", comment.Text);
        Assert.Equal(new DateTimeOffset(2026, 11, 24, 10, 0, 0, TimeSpan.Zero), comment.Date);

        document.Properties.Title = "Ny titel";
        document.Properties.Created = new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.Zero);
        OrdDocument saved = await TestFiles.OpenAsync(await TestFiles.SaveAsync(document));
        Assert.Equal("Ny titel", saved.Properties.Title);
        Assert.Equal(new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.Zero), saved.Properties.Created);
    }

    internal const string TrackedBody = """
        <w:p>
          <w:r><w:t xml:space="preserve">Det var en </w:t></w:r>
          <w:del w:id="1" w:author="A" w:date="2026-11-24T00:00:00Z"><w:r><w:delText>gammal</w:delText></w:r></w:del>
          <w:ins w:id="2" w:author="A" w:date="2026-11-24T00:00:00Z"><w:r><w:t>ny</w:t></w:r></w:ins>
          <w:r><w:t xml:space="preserve"> text.</w:t></w:r>
        </w:p>
        """;

    internal static async Task<byte[]> PartAsync(byte[] file, string name)
    {
        Result<Avig.OpenXml.OpcPackage> opened = await Avig.OpenXml.OpcPackage.OpenAsync(new MemoryStream(file));
        Assert.True(opened.IsSuccess(out Avig.OpenXml.OpcPackage? package));

        await using (package)
        {
            Assert.True((await package.ReadPartAsync(name)).IsSuccess(out byte[]? data), name);

            return data;
        }
    }

    private static async Task<byte[]> ReplaceInPartAsync(byte[] file, string part, string from, string to)
    {
        using MemoryStream copy = new();
        copy.Write(file);

        using (System.IO.Compression.ZipArchive archive = new(copy, System.IO.Compression.ZipArchiveMode.Update, true))
        {
            System.IO.Compression.ZipArchiveEntry entry = archive.GetEntry(part)!;
            string text;

            using (StreamReader reader = new(entry.Open()))
            {
                text = await reader.ReadToEndAsync();
            }

            entry.Delete();

            using StreamWriter writer = new(archive.CreateEntry(part).Open());
            await writer.WriteAsync(text.Replace(from, to, StringComparison.Ordinal));
        }

        return copy.ToArray();
    }

    private static Task<byte[]> ConvertToMacroAsync(byte[] file)
    {
        using MemoryStream copy = new();
        copy.Write(file);

        using (WordprocessingDocument document = WordprocessingDocument.Open(copy, true))
        {
            document.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.MacroEnabledDocument);
            VbaProjectPart vba = document.MainDocumentPart!.AddNewPart<VbaProjectPart>();

            using Stream stream = vba.GetStream(FileMode.Create);
            stream.Write([0xCC, 0x61, 0xFF, 0xFF]);
        }

        return Task.FromResult(copy.ToArray());
    }
}
