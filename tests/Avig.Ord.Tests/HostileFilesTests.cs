using System.Buffers.Binary;
using System.Diagnostics;
using System.IO.Compression;
using System.Text;

namespace Avig.Ord.Tests;

/// <summary>
/// Files made to hurt a reader, from the security test plan (docs/testplan.md), built in memory.
/// Each must be refused with a Result failure that has the right <see cref="OrdErrorKind"/>, or
/// read without crashing, within bounded memory and time. Nothing may be read from outside the
/// file. The class does not run beside other tests, so that its memory measurements are its own.
/// </summary>
[Collection(HostileFilesGroup.Name)]
public class HostileFilesTests
{
    private const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string Billion = """<!DOCTYPE x [<!ENTITY a "aaaaaaaaaa"><!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;"><!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;"><!ENTITY d "&c;&c;&c;&c;&c;&c;&c;&c;&c;&c;">]>""";

    // ---- Zip bombs ---------------------------------------------------------------------------

    [Fact]
    public async Task A_zip_bomb_past_the_limit_is_too_large_and_is_not_unpacked()
    {
        // 600 MB of zeros, which deflate to well under a megabyte; the default limit is 512 MiB.
        byte[] file = HostilePackage.Zip([.. HostilePackage.WordParts("<w:p/>"), ("word/media/bomb.bin", Zeros(600_000_000))], CompressionLevel.SmallestSize);
        long allocated = GC.GetTotalAllocatedBytes(true);
        Stopwatch watch = Stopwatch.StartNew();

        Result<OrdDocument> opened = await OrdDocument.OpenAsync(new MemoryStream(file));
        Result<DocxReader> read = await DocxReader.OpenAsync(new MemoryStream(file));

        Assert.Equal(OrdErrorKind.TooLarge, opened.ErrorKind());
        Assert.Equal(OrdErrorKind.TooLarge, read.ErrorKind());
        Assert.True(file.Length < 2_000_000, $"The bomb is {file.Length:N0} bytes.");
        Assert.True(GC.GetTotalAllocatedBytes(true) - allocated < 200_000_000, "The bomb was unpacked into memory.");
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(30), $"Refusing the bomb took {watch.Elapsed}.");
    }

    [Theory]
    [InlineData(100u)]
    [InlineData(0u)]
    [InlineData(0xFFFF_FFFFu)]
    public async Task A_part_whose_headers_lie_about_its_size_is_read_no_further_than_they_say_or_refused(uint statedSize)
    {
        // A 2 MB part whose local and central headers say it unpacks to something else. The zip
        // reader reads as much as the headers say and no more, so a small stated size reads as a
        // short part within the limit, and a large one is refused before anything is unpacked.
        byte[] file = HostilePackage.Zip([.. HostilePackage.WordParts("<w:p/>"), ("word/media/lie.bin", Zeros(2_000_000))], CompressionLevel.SmallestSize);
        HostilePackage.SetUncompressedSize(file, "word/media/lie.bin", statedSize);
        ReadOptions options = new() { Limits = OrdLimits.Default with { MaxUncompressedBytes = 1_000_000 } };

        long allocated = GC.GetTotalAllocatedBytes(true);

        Result<OrdDocument> opened = await OrdDocument.OpenAsync(new MemoryStream(file), options);

        if (statedSize > 1_000_000)
        {
            Assert.Equal(OrdErrorKind.TooLarge, opened.ErrorKind());
        }
        else if (opened.IsSuccess(out OrdDocument? document) == true)
        {
            Assert.Equal(string.Empty, document.ToPlainText());
            byte[] saved = await TestFiles.SaveAsync(document);
            Assert.True(saved.Length < 100_000, $"The saved file is {saved.Length:N0} bytes: more was unpacked than the headers say.");
        }
        else
        {
            Assert.Equal(OrdErrorKind.Damaged, opened.ErrorKind());
        }

        Assert.True(GC.GetTotalAllocatedBytes(true) - allocated < 50_000_000, "The part was unpacked past its stated size.");
    }

    [Fact]
    public async Task A_file_with_more_parts_than_the_limit_is_too_large()
    {
        List<(string Name, byte[] Data)> parts = [.. HostilePackage.WordParts("<w:p/>")];

        for (int index = 0; index < 10_001; index++)
        {
            parts.Add(($"word/x/{index}.bin", [1]));
        }

        byte[] file = HostilePackage.Zip(parts, CompressionLevel.Fastest);

        Assert.Equal(OrdErrorKind.TooLarge, (await OrdDocument.OpenAsync(new MemoryStream(file))).ErrorKind());
        Assert.Equal(OrdErrorKind.TooLarge, (await DocxReader.OpenAsync(new MemoryStream(file))).ErrorKind());
    }

    // ---- Nesting -----------------------------------------------------------------------------

    [Fact]
    public async Task Tables_in_tables_past_the_depth_limit_are_too_large()
    {
        byte[] file = HostilePackage.Docx(NestedTables(90));

        Assert.Equal(OrdErrorKind.TooLarge, (await OrdDocument.OpenAsync(new MemoryStream(file))).ErrorKind());

        Result<DocxReader> opened = await DocxReader.OpenAsync(new MemoryStream(file));
        Assert.True(opened.IsSuccess(out DocxReader? reader));

        await using (reader)
        {
            Exception thrown = await Assert.ThrowsAnyAsync<Exception>(async () =>
            {
                await foreach (Block block in reader.ReadBlocksAsync())
                {
                    _ = block.Text;
                }
            });

            Assert.Equal(OrdErrorKind.TooLarge, thrown.ErrorKind());
        }
    }

    [Fact]
    public async Task Tables_in_tables_just_inside_the_depth_limit_are_read_without_running_out_of_stack()
    {
        byte[] file = HostilePackage.Docx(NestedTables(80));

        OrdDocument document = await TestFiles.OpenAsync(file);

        // ToPlainText writes tables 8 deep and leaves deeper ones out, by design; the model and the other readers have the text.
        _ = document.ToPlainText();
        Assert.Contains("innerst", document.ToMarkdown(), StringComparison.Ordinal);
        Assert.Equal("innerst", Assert.Single(document.Paragraphs, paragraph => paragraph.Text.Length > 0).Text);
        Assert.Equal("innerst", (await ReadEventTextAsync(file)).Trim());
        Assert.Equal(string.Empty, TestFiles.Validate(await TestFiles.SaveAsync(document)).Split('\n').FirstOrDefault(line => line.Contains("nested", StringComparison.Ordinal)) ?? string.Empty);
    }

    [Fact]
    public async Task Hyperlinks_in_hyperlinks_and_text_boxes_in_text_boxes_are_read_once()
    {
        string links = string.Concat(Enumerable.Repeat("<w:hyperlink w:anchor=\"a\">", 60)) + "<w:r><w:t>länk</w:t></w:r>" + string.Concat(Enumerable.Repeat("</w:hyperlink>", 60));
        string boxes = "<w:r><w:t>yttre</w:t></w:r>";

        for (int depth = 0; depth < 12; depth++)
        {
            boxes = $"<w:r><w:drawing><w:txbxContent><w:p>{boxes}</w:p></w:txbxContent></w:drawing></w:r><w:r><w:t>{depth}</w:t></w:r>";
        }

        OrdDocument document = await TestFiles.OpenAsync(HostilePackage.Docx($"<w:p>{links}</w:p><w:p>{boxes}</w:p>"));

        Assert.StartsWith("länk", document.ToPlainText(), StringComparison.Ordinal);
        Assert.Equal(1, CountOf(document.ToPlainText(), "länk"));
        Assert.True(CountOf(document.ToPlainText(), "yttre") <= 1, document.ToPlainText());
    }

    // ---- DTDs --------------------------------------------------------------------------------

    [Theory]
    [InlineData("word/document.xml")]
    [InlineData("word/styles.xml")]
    [InlineData("word/numbering.xml")]
    [InlineData("word/header1.xml")]
    [InlineData("word/footnotes.xml")]
    [InlineData("word/comments.xml")]
    [InlineData("word/settings.xml")]
    [InlineData("docProps/core.xml")]
    [InlineData("word/_rels/document.xml.rels")]
    [InlineData("_rels/.rels")]
    [InlineData("[Content_Types].xml")]
    public async Task A_dtd_in_any_xml_part_is_refused_and_its_entities_are_never_expanded(string part)
    {
        List<(string Name, byte[] Data)> parts = HostilePackage.FullParts("<w:p><w:r><w:t>&amp;</w:t></w:r></w:p>");
        int index = parts.FindIndex(entry => entry.Name == part);
        string xml = Encoding.UTF8.GetString(parts[index].Data);
        int declaration = 0;

        if (xml.StartsWith("<?xml", StringComparison.Ordinal) == true)
        {
            declaration = xml.IndexOf("?>", StringComparison.Ordinal);
        }

        int root = xml.IndexOf('<', declaration);
        string rootName = xml[(root + 1)..xml.IndexOfAny([' ', '>', '/'], root + 1)];
        parts[index] = (part, Encoding.UTF8.GetBytes(xml.Insert(root, Billion.Replace("<!DOCTYPE x", "<!DOCTYPE " + rootName, StringComparison.Ordinal)).Replace("&amp;", "&d;", StringComparison.Ordinal)));
        byte[] file = HostilePackage.Zip(parts, CompressionLevel.Fastest);
        Stopwatch watch = Stopwatch.StartNew();

        Result<OrdDocument> opened = await OrdDocument.OpenAsync(new MemoryStream(file));
        Result<DocxReader> read = await DocxReader.OpenAsync(new MemoryStream(file));

        Assert.False(opened.IsSuccess(out _), part);
        Assert.NotEqual(OrdErrorKind.Unknown, opened.ErrorKind());

        if (read.IsSuccess(out DocxReader? reader) == true)
        {
            // The streaming reader reads the body and the styles as it goes; a DTD elsewhere is never parsed.
            await using (reader)
            {
                try
                {
                    _ = await reader.ReadPlainTextAsync();
                }
                catch (System.Xml.XmlException exception)
                {
                    Assert.Equal(OrdErrorKind.Damaged, exception.ErrorKind());
                }
            }
        }
        else
        {
            Assert.NotEqual(OrdErrorKind.Unknown, read.ErrorKind());
        }

        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(10), $"{part}: {watch.Elapsed}");
    }

    // ---- Styles and numbering ----------------------------------------------------------------

    [Fact]
    public async Task Styles_based_on_each_other_in_a_loop_and_in_a_very_long_chain_are_read()
    {
        StringBuilder styles = new();
        styles.Append("<w:style w:type=\"paragraph\" w:styleId=\"A\"><w:basedOn w:val=\"B\"/><w:rPr><w:b/></w:rPr></w:style>");
        styles.Append("<w:style w:type=\"paragraph\" w:styleId=\"B\"><w:basedOn w:val=\"A\"/><w:rPr><w:i/></w:rPr></w:style>");
        styles.Append("<w:style w:type=\"character\" w:styleId=\"Self\"><w:basedOn w:val=\"Self\"/></w:style>");

        for (int index = 0; index < 5_000; index++)
        {
            styles.Append(System.Globalization.CultureInfo.InvariantCulture, $"<w:style w:type=\"paragraph\" w:styleId=\"S{index}\"><w:basedOn w:val=\"S{index + 1}\"/><w:pPr><w:outlineLvl w:val=\"{index % 9}\"/></w:pPr></w:style>");
        }

        byte[] file = TestFiles.Docx(
            "<w:p><w:pPr><w:pStyle w:val=\"A\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"Self\"/></w:rPr><w:t>loop</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val=\"S0\"/></w:pPr><w:r><w:t>chain</w:t></w:r></w:p>",
            styles.ToString());
        Stopwatch watch = Stopwatch.StartNew();

        OrdDocument document = await TestFiles.OpenAsync(file);

        Assert.True(document.Paragraphs[0].Runs[0].Format.Bold);
        Assert.NotNull(document.Paragraphs[1].OutlineLevel);
        Assert.Equal("loop\nchain", document.ToPlainText());
        Assert.Contains("chain", document.ToMarkdown(), StringComparison.Ordinal);
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(10), watch.Elapsed.ToString());
    }

    [Theory]
    [InlineData("lowerLetter", "2147483647")]
    [InlineData("upperLetter", "100000000")]
    [InlineData("lowerRoman", "2147483647")]
    [InlineData("decimalZero", "-2147483648")]
    [InlineData("bullet", "2147483647")]
    [InlineData("decimal", "99999999999999999999")]
    public async Task List_definitions_with_absurd_levels_and_starts_give_short_labels(string format, string start)
    {
        string numbering = $"""
            <w:abstractNum w:abstractNumId="0">
              <w:lvl w:ilvl="0"><w:start w:val="{start}"/><w:numFmt w:val="{format}"/><w:lvlText w:val="%1.%2.%9.%0%"/></w:lvl>
              <w:lvl w:ilvl="99"><w:start w:val="1"/></w:lvl><w:lvl w:ilvl="-1"><w:start w:val="1"/></w:lvl><w:lvl><w:start w:val="1"/></w:lvl>
            </w:abstractNum>
            <w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="{start}"/></w:lvlOverride><w:lvlOverride w:ilvl="42"><w:startOverride w:val="1"/></w:lvlOverride></w:num>
            <w:num w:numId="2"><w:abstractNumId w:val="999"/></w:num>
            """;
        string item(string level, string id) => $"<w:p><w:pPr><w:numPr><w:ilvl w:val=\"{level}\"/><w:numId w:val=\"{id}\"/></w:numPr></w:pPr><w:r><w:t>punkt</w:t></w:r></w:p>";
        byte[] file = TestFiles.Docx(string.Concat(Enumerable.Repeat(item("0", "1") + item("99", "1") + item("-5", "1") + item("0", "2"), 50)), null, numbering);
        long allocated = GC.GetTotalAllocatedBytes(true);

        OrdDocument document = await TestFiles.OpenAsync(file);
        string text = document.ToPlainText();
        string markdown = document.ToMarkdown();

        Assert.All(document.Paragraphs, paragraph => Assert.True((paragraph.List?.Label.Length ?? 0) < 200, paragraph.List?.Label.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        Assert.True(text.Length < 100_000, $"The text is {text.Length:N0} characters.");
        Assert.True(markdown.Length < 100_000, $"The Markdown is {markdown.Length:N0} characters.");
        Assert.True(GC.GetTotalAllocatedBytes(true) - allocated < 200_000_000, "The labels took too much memory.");
    }

    // ---- Relationships -----------------------------------------------------------------------

    [Theory]
    [InlineData("../../../../Windows/win.ini")]
    [InlineData("/../secret.xml")]
    [InlineData("C:\\Windows\\win.ini")]
    [InlineData("file:///C:/Windows/win.ini")]
    [InlineData("\\\\server\\share\\x.xml")]
    [InlineData("%2e%2e/%2e%2e/x.xml")]
    [InlineData("media/missing.png")]
    [InlineData("")]
    public async Task Relationships_that_climb_out_of_the_package_or_point_to_nothing_read_as_missing(string target)
    {
        string relationships = $"""
            <Relationship Id="rIdHeader" Type="{Relationships}/header" Target="{target}"/>
            <Relationship Id="rIdImage" Type="{Relationships}/image" Target="{target}"/>
            <Relationship Id="rIdStyles" Type="{Relationships}/styles" Target="{target}"/>
            <Relationship Id="rIdNumbering" Type="{Relationships}/numbering" Target="{target}"/>
            <Relationship Id="rIdNotes" Type="{Relationships}/footnotes" Target="{target}"/>
            """;
        string body = "<w:p><w:r><w:t>text</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r>" + Picture("rIdImage") + "</w:p><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/></w:sectPr>";
        byte[] file = HostilePackage.Docx(body, relationships);

        Result<OrdDocument> opened = await OrdDocument.OpenAsync(new MemoryStream(file));

        if (opened.IsSuccess(out OrdDocument? document) == false)
        {
            Assert.Equal(OrdErrorKind.Damaged, opened.ErrorKind());

            return;
        }

        Assert.Equal("text", document.ToPlainText());
        Assert.Empty(document.Headers);
        Assert.Empty(document.Footnotes);
        Picture picture = Assert.Single(document.Paragraphs[0].Inlines.OfType<Picture>());
        Assert.Equal(ResultStatus.NotFound, (await picture.ReadAsync()).Status);
        Assert.Contains("text", document.ToMarkdown(), StringComparison.Ordinal);
        Assert.Equal("text", (await TestFiles.OpenAsync(await TestFiles.SaveAsync(document))).ToPlainText());
    }

    [Fact]
    public async Task A_picture_whose_part_is_missing_is_read_as_missing()
    {
        byte[] file = HostilePackage.Docx("<w:p>" + Picture("rIdImage") + "</w:p>", $"<Relationship Id=\"rIdImage\" Type=\"{Relationships}/image\" Target=\"media/image1.png\"/>");

        OrdDocument document = await TestFiles.OpenAsync(file);
        Picture picture = Assert.Single(document.Paragraphs[0].Inlines.OfType<Picture>());

        Assert.Equal("word/media/image1.png", picture.PartName);
        Assert.Equal(ResultStatus.NotFound, (await picture.ReadAsync()).Status);
        Assert.Contains("![bild](media/image1.png)", document.ToMarkdown(), StringComparison.Ordinal);
        await TestFiles.SaveAsync(document);
    }

    // ---- Fields and symbols ------------------------------------------------------------------

    [Theory]
    [InlineData("<w:r><w:t>före</w:t></w:r><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:t>efter</w:t></w:r>", "före")]
    [InlineData("<w:r><w:t>före</w:t></w:r><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>4</w:t></w:r>", "före4")]
    [InlineData("<w:r><w:t>före</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>efter</w:t></w:r>", "föreefter")]
    [InlineData("<w:r><w:fldChar w:fldCharType=\"end\"/><w:fldChar w:fldCharType=\"end\"/><w:fldChar w:fldCharType=\"bogus\"/><w:fldChar/><w:t>x</w:t></w:r>", "x")]
    public async Task A_field_begin_without_end_and_an_end_without_begin_are_read(string paragraph, string expected)
    {
        byte[] file = HostilePackage.Docx($"<w:p>{paragraph}</w:p><w:p><w:r><w:t>nästa</w:t></w:r></w:p>");

        OrdDocument document = await TestFiles.OpenAsync(file);

        Assert.Equal(expected + "\nnästa", document.ToPlainText());
        Assert.Equal(expected + "\nnästa", await ReadEventTextAsync(file));
        _ = document.ToMarkdown();
    }

    [Fact]
    public async Task Many_fields_begun_and_never_ended_are_read_in_bounded_time()
    {
        string begins = string.Concat(Enumerable.Repeat("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:hyperlink w:anchor=\"a\"><w:r><w:t>x</w:t></w:r>", 20_000))
            + string.Concat(Enumerable.Repeat("</w:hyperlink>", 20_000));
        Stopwatch watch = Stopwatch.StartNew();

        Result<OrdDocument> opened = await OrdDocument.OpenAsync(new MemoryStream(HostilePackage.Docx($"<w:p>{begins}</w:p>")));

        // Nested 20 000 deep, this is past the depth limit; a flat run of them is not.
        Assert.Equal(OrdErrorKind.TooLarge, opened.ErrorKind());

        string flat = string.Concat(Enumerable.Repeat("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>x</w:t></w:r>", 20_000));
        OrdDocument document = await TestFiles.OpenAsync(HostilePackage.Docx($"<w:p>{flat}</w:p>"));

        Assert.Equal(20_000, document.ToPlainText().Length);
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(20), watch.Elapsed.ToString());
    }

    [Theory]
    [InlineData("FFFFFFFF")]
    [InlineData("110000")]
    [InlineData("D800")]
    [InlineData("DFFF")]
    [InlineData("-1")]
    [InlineData("0")]
    [InlineData("99999999999999999999999")]
    [InlineData("zz")]
    [InlineData("")]
    public async Task A_symbol_that_is_no_character_gives_no_text(string code)
    {
        byte[] file = HostilePackage.Docx($"<w:p><w:r><w:t>a</w:t><w:sym w:font=\"Wingdings\" w:char=\"{code}\"/><w:t>b</w:t></w:r></w:p>");

        OrdDocument document = await TestFiles.OpenAsync(file);

        Assert.Equal("ab", document.ToPlainText());
        Assert.Equal("ab", await ReadEventTextAsync(file));
    }

    // ---- Flat OPC ----------------------------------------------------------------------------

    [Theory]
    [InlineData("duplicate", OrdErrorKind.Damaged)]
    [InlineData("outside", OrdErrorKind.Damaged)]
    [InlineData("base64", OrdErrorKind.Damaged)]
    [InlineData("two-roots", OrdErrorKind.Damaged)]
    [InlineData("dtd", OrdErrorKind.Damaged)]
    [InlineData("not-a-package", OrdErrorKind.NotWord)]
    [InlineData("spreadsheet", OrdErrorKind.NotWord)]
    [InlineData("no-main-part", OrdErrorKind.NotWord)]
    public async Task Flat_opc_files_made_to_hurt_are_refused(string variant, OrdErrorKind kind)
    {
        string document = $"<pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"{MainType(variant)}\"><pkg:xmlData><w:document xmlns:w=\"{TestFiles.Wml}\"><w:body><w:p/></w:body></w:document></pkg:xmlData></pkg:part>";
        string rels = $"<pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\"><pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{Relationships}/officeDocument\" Target=\"word/document.xml\"/></Relationships></pkg:xmlData></pkg:part>";
        string extra = variant switch
        {
            "duplicate" => document,
            "outside" => "<pkg:part pkg:name=\"/../../outside.xml\" pkg:contentType=\"application/xml\"><pkg:xmlData><x/></pkg:xmlData></pkg:part>",
            "base64" => "<pkg:part pkg:name=\"/word/media/image1.png\" pkg:contentType=\"image/png\" pkg:compression=\"store\"><pkg:binaryData>!!! not base64 !!!</pkg:binaryData></pkg:part>",
            "two-roots" => "<pkg:part pkg:name=\"/word/settings.xml\" pkg:contentType=\"application/xml\"><pkg:xmlData><a/><b/></pkg:xmlData></pkg:part>",
            _ => string.Empty,
        };

        if (variant == "no-main-part")
        {
            document = string.Empty;
        }

        string flat = $"<?xml version=\"1.0\" standalone=\"yes\"?><?mso-application progid=\"Word.Document\"?><pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">{rels}{document}{extra}</pkg:package>";

        if (variant == "dtd")
        {
            flat = flat.Replace("<pkg:package", Billion.Replace("<!DOCTYPE x", "<!DOCTYPE pkg:package", StringComparison.Ordinal) + "<pkg:package", StringComparison.Ordinal).Replace("<w:p/>", "<w:p><w:r><w:t>&d;</w:t></w:r></w:p>", StringComparison.Ordinal);
        }
        else if (variant == "not-a-package")
        {
            flat = "<?xml version=\"1.0\"?><html><body>not a package</body></html>";
        }

        byte[] file = Encoding.UTF8.GetBytes(flat);

        Assert.Equal(kind, (await OrdDocument.OpenAsync(new MemoryStream(file))).ErrorKind());
        Assert.Equal(kind, (await DocxReader.OpenAsync(new MemoryStream(file))).ErrorKind());
    }

    [Fact]
    public async Task A_docx_whose_main_part_is_a_spreadsheet_is_not_word()
    {
        const string sheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        List<(string Name, byte[] Data)> parts =
        [
            ("[Content_Types].xml", Encoding.UTF8.GetBytes($"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"{sheet}\"/></Types>")),
            ("_rels/.rels", Encoding.UTF8.GetBytes($"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{Relationships}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>")),
            ("xl/workbook.xml", "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheets/></workbook>"u8.ToArray()),
        ];
        byte[] file = HostilePackage.Zip(parts, CompressionLevel.Fastest);

        Assert.Equal(OrdErrorKind.NotWord, (await OrdDocument.OpenAsync(new MemoryStream(file))).ErrorKind());
        Assert.Equal(OrdErrorKind.NotWord, (await DocxReader.OpenAsync(new MemoryStream(file))).ErrorKind());

        // A spreadsheet's XML in a part that claims to be a Word document is read as an empty document.
        List<(string Name, byte[] Data)> claimed = HostilePackage.WordParts(string.Empty);
        claimed[claimed.FindIndex(part => part.Name == "word/document.xml")] = ("word/document.xml", "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row><c><v>1</v></c></row></sheetData></worksheet>"u8.ToArray());
        OrdDocument document = await TestFiles.OpenAsync(HostilePackage.Zip(claimed, CompressionLevel.Fastest));

        Assert.Equal(string.Empty, document.ToPlainText());
        document.AddParagraph("ny");
        Assert.Contains("ny", (await TestFiles.OpenAsync(await TestFiles.SaveAsync(document))).ToPlainText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(new byte[] { 0x50, 0x4B, 0x03, 0x04 })]
    [InlineData(new byte[] { 0x50, 0x4B, 0x05, 0x06, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 })]
    [InlineData(new byte[] { 0xEF, 0xBB, 0xBF, (byte)'<' })]
    [InlineData(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 })]
    [InlineData(new byte[0])]
    public async Task Truncated_and_empty_files_are_refused_with_a_kind(byte[] file)
    {
        Assert.NotEqual(OrdErrorKind.Unknown, (await OrdDocument.OpenAsync(new MemoryStream(file))).ErrorKind());
        Assert.NotEqual(OrdErrorKind.Unknown, (await DocxReader.OpenAsync(new MemoryStream(file))).ErrorKind());
    }

    // ---- Helpers -----------------------------------------------------------------------------

    private static string NestedTables(int depth)
    {
        return string.Concat(Enumerable.Repeat("<w:tbl><w:tr><w:tc>", depth)) + "<w:p><w:r><w:t>innerst</w:t></w:r></w:p>" + string.Concat(Enumerable.Repeat("</w:tc></w:tr></w:tbl>", depth));
    }

    private static string Picture(string id)
    {
        return $"<w:r><w:drawing><wp:inline xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"><wp:extent cx=\"9525\" cy=\"9525\"/><wp:docPr id=\"1\" name=\"bild\" descr=\"bild\"/><a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData><a:blip r:embed=\"{id}\"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
    }

    private static string MainType(string variant)
    {
        if (variant == "spreadsheet")
        {
            return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        }

        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    }

    private static int CountOf(string text, string word)
    {
        int count = 0;
        int position = 0;

        while ((position = text.IndexOf(word, position, StringComparison.Ordinal)) >= 0)
        {
            count++;
            position += word.Length;
        }

        return count;
    }

    private static byte[] Zeros(int length)
    {
        return new byte[length];
    }

    /// <summary>The body's text as the events give it: paragraphs on lines of their own.</summary>
    internal static async Task<string> ReadEventTextAsync(byte[] file)
    {
        Result<DocxReader> opened = await DocxReader.OpenAsync(new MemoryStream(file));
        Assert.True(opened.IsSuccess(out DocxReader? reader), opened.Error);
        List<string> paragraphs = [];
        StringBuilder current = new();

        await using (reader)
        {
            await foreach (DocumentEvent item in reader.ReadEventsAsync())
            {
                if (item.Kind == DocumentEventKind.ParagraphStart)
                {
                    current.Clear();
                }
                else if (item.Kind is DocumentEventKind.Text)
                {
                    current.Append(item.Text);
                }
                else if (item.Kind is DocumentEventKind.Break)
                {
                    current.Append('\n');
                }
                else if (item.Kind == DocumentEventKind.ParagraphEnd)
                {
                    paragraphs.Add(current.ToString());
                }
            }
        }

        return string.Join('\n', paragraphs);
    }
}

/// <summary>The hostile files measure memory; they run alone.</summary>
[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class HostileFilesGroup
{
    public const string Name = "Hostile files";
}

/// <summary>Word packages built part by part, for files no library would write.</summary>
internal static class HostilePackage
{
    private const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string Word = "application/vnd.openxmlformats-officedocument.wordprocessingml";

    /// <summary>A .docx with this body (r: and w: declared) and these extra relationships from the main part.</summary>
    public static byte[] Docx(string body, string relationships = "")
    {
        return Zip(WordParts(body, relationships), CompressionLevel.Fastest);
    }

    /// <summary>The smallest Word package: content types, the package relationships, the main part and its relationships.</summary>
    public static List<(string Name, byte[] Data)> WordParts(string body, string relationships = "")
    {
        return
        [
            ("[Content_Types].xml", Encoding.UTF8.GetBytes($"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"bin\" ContentType=\"application/octet-stream\"/><Override PartName=\"/word/document.xml\" ContentType=\"{Word}.document.main+xml\"/></Types>")),
            ("_rels/.rels", Encoding.UTF8.GetBytes($"<Relationships xmlns=\"{PackageRelationships}\"><Relationship Id=\"rId1\" Type=\"{Relationships}/officeDocument\" Target=\"word/document.xml\"/></Relationships>")),
            ("word/document.xml", Encoding.UTF8.GetBytes($"<w:document xmlns:w=\"{TestFiles.Wml}\" xmlns:r=\"{Relationships}\"><w:body>{body}</w:body></w:document>")),
            ("word/_rels/document.xml.rels", Encoding.UTF8.GetBytes($"<Relationships xmlns=\"{PackageRelationships}\">{relationships}</Relationships>")),
        ];
    }

    /// <summary>A package with every part the document model reads: styles, numbering, a header, footnotes, comments, settings and core properties.</summary>
    public static List<(string Name, byte[] Data)> FullParts(string body)
    {
        string overrides = string.Concat(
            $"<Override PartName=\"/word/document.xml\" ContentType=\"{Word}.document.main+xml\"/>",
            $"<Override PartName=\"/word/styles.xml\" ContentType=\"{Word}.styles+xml\"/>",
            $"<Override PartName=\"/word/numbering.xml\" ContentType=\"{Word}.numbering+xml\"/>",
            $"<Override PartName=\"/word/header1.xml\" ContentType=\"{Word}.header+xml\"/>",
            $"<Override PartName=\"/word/footnotes.xml\" ContentType=\"{Word}.footnotes+xml\"/>",
            $"<Override PartName=\"/word/comments.xml\" ContentType=\"{Word}.comments+xml\"/>",
            $"<Override PartName=\"/word/settings.xml\" ContentType=\"{Word}.settings+xml\"/>",
            "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
        string wordNamespace = $"xmlns:w=\"{TestFiles.Wml}\"";

        return
        [
            ("[Content_Types].xml", Encoding.UTF8.GetBytes($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>{overrides}</Types>")),
            ("_rels/.rels", Encoding.UTF8.GetBytes($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"{PackageRelationships}\"><Relationship Id=\"rId1\" Type=\"{Relationships}/officeDocument\" Target=\"word/document.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/></Relationships>")),
            ("word/document.xml", Encoding.UTF8.GetBytes($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:document {wordNamespace} xmlns:r=\"{Relationships}\"><w:body>{body}<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rId1\"/></w:sectPr></w:body></w:document>")),
            ("word/_rels/document.xml.rels", Encoding.UTF8.GetBytes($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"{PackageRelationships}\"><Relationship Id=\"rId1\" Type=\"{Relationships}/header\" Target=\"header1.xml\"/><Relationship Id=\"rId2\" Type=\"{Relationships}/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId3\" Type=\"{Relationships}/numbering\" Target=\"numbering.xml\"/><Relationship Id=\"rId4\" Type=\"{Relationships}/footnotes\" Target=\"footnotes.xml\"/><Relationship Id=\"rId5\" Type=\"{Relationships}/comments\" Target=\"comments.xml\"/><Relationship Id=\"rId6\" Type=\"{Relationships}/settings\" Target=\"settings.xml\"/></Relationships>")),
            ("word/styles.xml", Encoding.UTF8.GetBytes($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:styles {wordNamespace}><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"&amp;\"/></w:style></w:styles>")),
            ("word/numbering.xml", Encoding.UTF8.GetBytes($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:numbering {wordNamespace}><w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:lvlText w:val=\"&amp;\"/></w:lvl></w:abstractNum></w:numbering>")),
            ("word/header1.xml", Encoding.UTF8.GetBytes($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:hdr {wordNamespace}><w:p><w:r><w:t>&amp;</w:t></w:r></w:p></w:hdr>")),
            ("word/footnotes.xml", Encoding.UTF8.GetBytes($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:footnotes {wordNamespace}><w:footnote w:id=\"1\"><w:p><w:r><w:t>&amp;</w:t></w:r></w:p></w:footnote></w:footnotes>")),
            ("word/comments.xml", Encoding.UTF8.GetBytes($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:comments {wordNamespace}><w:comment w:id=\"0\"><w:p><w:r><w:t>&amp;</w:t></w:r></w:p></w:comment></w:comments>")),
            ("word/settings.xml", Encoding.UTF8.GetBytes($"<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:settings {wordNamespace}><w:zoom w:percent=\"100\"/><w:rsids><w:rsidRoot w:val=\"00\"/></w:rsids><w:docVars><w:docVar w:name=\"x\" w:val=\"&amp;\"/></w:docVars></w:settings>")),
            ("docProps/core.xml", Encoding.UTF8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>&amp;</dc:title></cp:coreProperties>")),
        ];
    }

    public static byte[] Zip(IEnumerable<(string Name, byte[] Data)> parts, CompressionLevel level)
    {
        using MemoryStream file = new();

        using (ZipArchive archive = new(file, ZipArchiveMode.Create, true))
        {
            foreach ((string name, byte[] data) in parts)
            {
                using Stream entry = archive.CreateEntry(name, level).Open();
                entry.Write(data);
            }
        }

        return file.ToArray();
    }

    /// <summary>Writes another uncompressed size into the local and the central header of an entry.</summary>
    public static void SetUncompressedSize(byte[] file, string name, uint size)
    {
        byte[] encoded = Encoding.UTF8.GetBytes(name);

        for (int position = 0; position < file.Length - 46; position++)
        {
            bool local = file[position] == 0x50 && file[position + 1] == 0x4B && file[position + 2] == 3 && file[position + 3] == 4;
            bool central = file[position] == 0x50 && file[position + 1] == 0x4B && file[position + 2] == 1 && file[position + 3] == 2;

            if (local == true && file.AsSpan(position + 30).StartsWith(encoded) == true)
            {
                BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(position + 22), size);
            }
            else if (central == true && file.AsSpan(position + 46).StartsWith(encoded) == true)
            {
                BinaryPrimitives.WriteUInt32LittleEndian(file.AsSpan(position + 24), size);
            }
        }
    }
}
