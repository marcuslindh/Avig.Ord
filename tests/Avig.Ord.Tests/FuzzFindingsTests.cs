using System.Text;
using Avig.Ord.Internal;
using DocumentFormat.OpenXml.Packaging;

namespace Avig.Ord.Tests;

/// <summary>
/// Regression tests for what tools/Avig.Ord.Fuzz found (its README has the table), and for the
/// one bug the Word check found. Each test is the smallest input that showed the bug.
/// </summary>
public class FuzzFindingsTests
{
    private const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // ---- docx-part, zip-raw, roundtrip: OpenAsync succeeded and a later call threw ----------

    [Theory]
    [InlineData("<w:hdr xmlns:w=\"" + TestFiles.Wml + "\"><w:p><w:r><w:t>x</w:te></w:r></w:p></w:hdr>")]
    [InlineData("")]
    public async Task A_broken_or_empty_header_is_found_when_the_file_is_opened(string header)
    {
        byte[] file = TestFiles.Docx("<w:p/>", extra: main =>
        {
            HeaderPart part = main.AddNewPart<HeaderPart>("rIdHeader");
            TestFiles.Write(part, header);
        });

        Result<OrdDocument> opened = await OrdDocument.OpenAsync(new MemoryStream(file));

        // Before: OpenAsync succeeded, and Headers threw XmlException.
        Assert.Equal(OrdErrorKind.Damaged, opened.ErrorKind());
    }

    [Fact]
    public async Task Broken_xml_in_a_part_the_model_never_reads_is_found_when_the_file_is_opened()
    {
        byte[] file = TestFiles.Docx("<w:p/>", extra: main =>
        {
            CustomXmlPart custom = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            TestFiles.Write(custom, "<data><unclosed></data>");
        });

        // Before: OpenAsync succeeded, and SaveAsync as Flat OPC threw XmlException, since Flat OPC holds every XML part as XML.
        Assert.Equal(OrdErrorKind.Damaged, (await OrdDocument.OpenAsync(new MemoryStream(file))).ErrorKind());
    }

    [Fact]
    public async Task Two_relationships_with_one_id_are_a_damaged_file()
    {
        string relationships = $"<Relationship Id=\"rId1\" Type=\"{Relationships}/hyperlink\" Target=\"https://a.example/\" TargetMode=\"External\"/><Relationship Id=\"rId1\" Type=\"{Relationships}/hyperlink\" Target=\"https://b.example/\" TargetMode=\"External\"/>";
        byte[] file = HostilePackage.Docx("<w:p><w:hyperlink r:id=\"rId1\"><w:r><w:t>länk</w:t></w:r></w:hyperlink></w:p>", relationships);

        // Before: the file opened, and SaveAsync threw InvalidOperationException from the package writer.
        Assert.Equal(OrdErrorKind.Damaged, (await OrdDocument.OpenAsync(new MemoryStream(file))).ErrorKind());
    }

    [Fact]
    public async Task Flat_opc_xml_data_with_two_root_elements_is_a_damaged_file()
    {
        string flat = $"<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"><pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\"><pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{Relationships}/officeDocument\" Target=\"word/document.xml\"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"><pkg:xmlData><w:document xmlns:w=\"{TestFiles.Wml}\"/><w:document xmlns:w=\"{TestFiles.Wml}\"/></pkg:xmlData></pkg:part></pkg:package>";

        // Before: OpenAsync threw InvalidOperationException from inside Avig.OpenXml 0.1.2 (worked around in WordPackage).
        Assert.Equal(OrdErrorKind.Damaged, (await OrdDocument.OpenAsync(new MemoryStream(Encoding.UTF8.GetBytes(flat)))).ErrorKind());
        Assert.Equal(OrdErrorKind.Damaged, (await DocxReader.OpenAsync(new MemoryStream(Encoding.UTF8.GetBytes(flat)))).ErrorKind());
    }

    [Fact(Skip = "A bug in Avig.OpenXml 0.1.2, to be fixed there: OpcPackageWriter re-reads every XML part for Flat OPC with its own depth limit of 100.")]
    public async Task A_document_nested_deeper_than_100_levels_saves_as_flat_opc()
    {
        string body = string.Concat(Enumerable.Repeat("<w:tbl><w:tr><w:tc>", 40)) + "<w:p/>" + string.Concat(Enumerable.Repeat("</w:tc></w:tr></w:tbl>", 40));
        OrdDocument document = await TestFiles.OpenAsync(HostilePackage.Docx(body));

        byte[] flat = await TestFiles.SaveAsync(document, new SaveOptions { FlatOpc = true });

        Assert.Equal(document.ToPlainText(), (await TestFiles.OpenAsync(flat)).ToPlainText());
    }

    // ---- render: a field that ends inside a hyperlink ----------------------------------------

    [Fact]
    public async Task A_field_that_begins_before_a_hyperlink_and_ends_inside_it_keeps_the_text_in_order()
    {
        const string paragraph = """
            <w:p>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>a</w:t></w:r>
              <w:hyperlink w:anchor="x"><w:r><w:t>b</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>c</w:t></w:r></w:hyperlink>
              <w:r><w:t>d</w:t></w:r>
            </w:p>
            """;
        byte[] file = TestFiles.Docx(paragraph);

        OrdDocument document = await TestFiles.OpenAsync(file);

        // Before: InvalidOperationException "Stack empty", and the rest of the link came before its start.
        Assert.Equal("abcd", document.Paragraphs[0].Text);
        Assert.Equal("abcd", await HostileFilesTests.ReadEventTextAsync(file));
        Assert.Contains("[", document.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_field_begun_inside_a_hyperlink_and_not_ended_there_ends_with_it_in_both_readers()
    {
        const string paragraph = """
            <w:p>
              <w:hyperlink w:anchor="x"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>IF </w:instrText></w:r></w:hyperlink>
              <w:r><w:t>efter</w:t></w:r>
            </w:p>
            """;
        byte[] file = TestFiles.Docx(paragraph);

        Assert.Equal("efter", (await TestFiles.OpenAsync(file)).Paragraphs[0].Text);
        Assert.Equal("efter", await HostileFilesTests.ReadEventTextAsync(file));
    }

    // ---- reader-model: the streaming readers and the document model differed ----------------

    [Fact]
    public async Task The_event_reader_has_the_text_of_the_paragraphs()
    {
        const string body = """
            <w:tbl><w:tr><w:trPr><w:del w:id="1" w:author="A"/></w:trPr><w:tc><w:p><w:r><w:t>raderad rad</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>rad</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
            <w:tbl/><w:tbl><w:tr/><w:tr><w:tc/></w:tr></w:tbl>
            <w:p><w:r><w:t>ett</w:t><w:noBreakHyphen/><w:t>två</w:t><w:sym w:font="Wingdings" w:char="0041"/></w:r></w:p>
            <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>IF </w:instrText></w:r></w:p>
            <w:p><w:r><w:t>efter ett fält utan slut</w:t></w:r><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath></w:p>
            """;
        byte[] file = TestFiles.Docx(body);
        OrdDocument document = await TestFiles.OpenAsync(file);

        // Before: the deleted row was there, the empty table never ended, ‑ and the symbol were
        // missing, the field without an end hid the next paragraph, and the equation's x was text.
        Assert.Equal(string.Join('\n', document.Paragraphs.Select(paragraph => paragraph.Text)), await HostileFilesTests.ReadEventTextAsync(file));
        Assert.Equal(0, await TableDepthAtEndAsync(file));
    }

    [Fact]
    public async Task The_streaming_reader_gives_the_same_plain_text_and_markdown_as_the_model()
    {
        string item(string text) => $"<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";
        const string numbering = """<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>""";
        byte[] file = TestFiles.Docx(item("ett") + "<w:p/>" + item("två") + "<w:p><w:r><w:t>slut</w:t><w:br/></w:r></w:p><w:p/>", null, numbering);
        OrdDocument document = await TestFiles.OpenAsync(file);

        Result<DocxReader> forText = await DocxReader.OpenAsync(new MemoryStream(file));
        Assert.True(forText.IsSuccess(out DocxReader? textReader));
        Result<DocxReader> forMarkdown = await DocxReader.OpenAsync(new MemoryStream(file));
        Assert.True(forMarkdown.IsSuccess(out DocxReader? markdownReader));
        await using StringWriter markdown = new();

        await using (textReader)
        await using (markdownReader)
        {
            // Before: the streaming text lost every line end at the end, and the empty paragraph made two lists.
            Assert.Equal(document.ToPlainText(), await textReader.ReadPlainTextAsync());
            await markdownReader.WriteMarkdownAsync(markdown);
        }

        Assert.Equal(document.ToMarkdown(), markdown.ToString());
    }

    // ---- revisions: nested text boxes --------------------------------------------------------

    [Fact]
    public async Task A_text_box_in_a_text_box_is_read_once_and_in_the_view_of_its_row()
    {
        const string body = """
            <w:p><w:r><w:drawing><w:txbxContent>
              <w:p><w:r><w:t>yttre</w:t></w:r></w:p>
              <w:tbl><w:tr><w:trPr><w:ins w:id="1" w:author="A"/></w:trPr><w:tc><w:p><w:r><w:drawing><w:txbxContent><w:p><w:r><w:t>inre</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r></w:p></w:tc></w:tr></w:tbl>
            </w:txbxContent></w:drawing></w:r></w:p>
            """;
        byte[] file = TestFiles.Docx(body);

        OrdDocument accepted = await TestFiles.OpenAsync(file);
        OrdDocument original = await TestFiles.OpenAsync(file, new ReadOptions { Revisions = RevisionView.Original });
        OrdDocument rejected = await TestFiles.OpenAsync(file);
        rejected.RejectAllRevisions();

        // Before: the inner box was read as a box of the outer drawing too, twice in the text and
        // also in the Original view, where its row is not there.
        Assert.Equal(1, accepted.ToPlainText().Split("inre").Length - 1);
        Assert.DoesNotContain("inre", original.ToPlainText(), StringComparison.Ordinal);
        Assert.Equal(rejected.ToPlainText(), original.ToPlainText());
    }

    // ---- roundtrip: what the model wrote into a valid file was not valid ---------------------

    [Fact]
    public async Task A_table_added_to_a_page_whose_margins_are_wider_than_it_has_positive_column_widths()
    {
        byte[] file = TestFiles.Docx("<w:p/><w:sectPr><w:pgSz w:w=\"5000\" w:h=\"5000\"/><w:pgMar w:top=\"0\" w:right=\"4000\" w:bottom=\"0\" w:left=\"4000\" w:header=\"0\" w:footer=\"0\" w:gutter=\"0\"/></w:sectPr>");
        Assert.Equal(string.Empty, TestFiles.Validate(file));
        OrdDocument document = await TestFiles.OpenAsync(file);

        document.AddTable(table => table.Row("a", "b"));

        // Before: w:gridCol w:w="-1404".
        Assert.Equal(string.Empty, TestFiles.Validate(await TestFiles.SaveAsync(document)));
    }

    [Fact]
    public async Task New_lists_notes_and_pictures_get_valid_ids_when_the_file_already_uses_the_largest()
    {
        const string numbering = """<w:abstractNum w:abstractNumId="2147483647"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="2147483647"><w:abstractNumId w:val="2147483647"/></w:num>""";
        byte[] file = TestFiles.Docx("<w:p><w:r><w:t>text</w:t></w:r></w:p>", null, numbering);
        Assert.Equal(string.Empty, TestFiles.Validate(file));
        OrdDocument document = await TestFiles.OpenAsync(file);

        document.AddParagraph(paragraph => paragraph.Numbered().Text("ny lista"));
        document.AddParagraph(paragraph => paragraph.Bullet().Text("punkt"));

        // Before: abstractNumId and numId -2147483648.
        Assert.Equal(string.Empty, TestFiles.Validate(await TestFiles.SaveAsync(document)));
        Assert.Equal(1, OrdDocument.NextFreeId([0, int.MaxValue], 1));
        Assert.Equal(0, OrdDocument.NextFreeId([int.MaxValue], 0));
        Assert.Equal(6, OrdDocument.NextFreeId([1, 5], 1));
    }

    // ---- render: a gridSpan of int.MaxValue --------------------------------------------------

    [Fact]
    public async Task Spans_and_grids_wider_than_word_allows_neither_overflow_nor_make_huge_markdown()
    {
        string grid = string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"10\"/>", 100_000));
        string rows = string.Concat(Enumerable.Repeat("<w:tr><w:tc><w:p/></w:tc></w:tr>", 200));
        string body = $"<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val=\"2147483647\"/></w:tcPr><w:p/></w:tc><w:tc><w:tcPr><w:gridSpan w:val=\"2147483647\"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:tbl><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>";
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx(body));

        // Before: OverflowException from Enumerable.Sum, and 200 rows of 100 000 cells of Markdown.
        Assert.Equal(TableCell.MaxColumns, document.Tables[0].Rows[0].Cells[0].ColumnSpan);
        Assert.Equal(2 * TableCell.MaxColumns, document.Tables[0].ColumnCount);
        Assert.True(document.ToMarkdown().Length < 100_000);
    }

    // ---- image-size, format ------------------------------------------------------------------

    [Fact]
    public async Task A_bitmap_whose_height_is_int_min_value_has_no_size()
    {
        byte[] bitmap = new byte[30];
        "BM"u8.CopyTo(bitmap);
        BitConverter.GetBytes(10).CopyTo(bitmap, 18);
        BitConverter.GetBytes(int.MinValue).CopyTo(bitmap, 22);

        // Before: OverflowException from Math.Abs, when such a picture was added to a document;
        // now the ArgumentException the builder documents for a picture without a size.
        Assert.False(ImageInfo.TryGetSize(bitmap, out _, out _));

        await using DocxWriter writer = await DocxWriter.CreateAsync(new MemoryStream());
        await Assert.ThrowsAsync<ArgumentException>(() => writer.PictureAsync(bitmap));
        await writer.ParagraphAsync(paragraph => paragraph.Picture(bitmap, Length.Px(10), Length.Px(10)));
    }

    [Theory]
    [InlineData("20000000000000000")]
    [InlineData("-20000000000000000")]
    [InlineData("9223372036854775807")]
    [InlineData("-9223372036854775808")]
    public void A_length_too_large_to_hold_is_no_length_and_never_one_of_the_other_sign(string twips)
    {
        // Before: 20000000000000000 twips came back as a negative length.
        Assert.Null(FormatReader.Twips(twips));
        Assert.Equal(Length.FromTwips(1440), FormatReader.Twips("1440"));
    }

    // ---- Word check --------------------------------------------------------------------------

    [Fact]
    public async Task Notes_are_those_the_body_refers_to_in_its_order_as_word_lists_them()
    {
        byte[] file = TestFiles.Docx(
            "<w:p><w:r><w:t>a</w:t></w:r><w:r><w:footnoteReference w:id=\"3\"/></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>",
            extra: main =>
            {
                FootnotesPart notes = main.AddNewPart<FootnotesPart>();
                TestFiles.Write(notes, $"<w:footnotes xmlns:w=\"{TestFiles.Wml}\"><w:footnote w:type=\"separator\" w:id=\"-1\"><w:p/></w:footnote><w:footnote w:id=\"1\"><w:p><w:r><w:t>ett</w:t></w:r></w:p></w:footnote><w:footnote w:id=\"2\"><w:p><w:r><w:t>föräldralös</w:t></w:r></w:p></w:footnote><w:footnote w:id=\"3\"><w:p><w:r><w:t>tre</w:t></w:r></w:p></w:footnote></w:footnotes>");
            });

        OrdDocument document = await TestFiles.OpenAsync(file);

        // Before: Footnotes listed note 2, which nothing refers to (Paragraph.SetText had replaced
        // its reference), while Word showed no such note.
        Assert.Equal(["tre", "ett"], document.Footnotes.Select(note => note.Text));
    }

    private static async Task<int> TableDepthAtEndAsync(byte[] file)
    {
        Result<DocxReader> opened = await DocxReader.OpenAsync(new MemoryStream(file));
        Assert.True(opened.IsSuccess(out DocxReader? reader));
        int depth = 0;

        await using (reader)
        {
            await foreach (DocumentEvent item in reader.ReadEventsAsync())
            {
                if (item.Kind == DocumentEventKind.TableStart)
                {
                    depth++;
                }
                else if (item.Kind == DocumentEventKind.TableEnd)
                {
                    depth--;
                }
            }
        }

        return depth;
    }
}
