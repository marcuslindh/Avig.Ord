using System.Text;

namespace Avig.Ord.Fuzz;

/// <summary>
/// <c>docx-part</c> — a damaged or unusual part must give a Result failure with a kind, or the
/// exceptions <see cref="DocxReader"/> documents, never anything else; and what the document
/// model saves must open again with the change in it.
///
/// The input is not a whole .docx. Its first byte picks one part of the rich template (a document
/// <see cref="DocxWriter"/> wrote, see <see cref="Template"/>), the second byte the revision view
/// and whether to save as Flat OPC, and the rest replaces the part. Mutating a zip file directly
/// gets nowhere: nearly every change breaks the deflate stream or the CRC, so the fuzzer would
/// only test the zip reader. This way the mutations land in the XML the package parses.
/// </summary>
internal static class PartTarget
{
    /// <summary>The parts worth replacing, in the order the first input byte picks them.</summary>
    internal static readonly string[] PartNames =
    [
        "word/document.xml",
        "word/styles.xml",
        "word/numbering.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/header1.xml",
        "word/footer1.xml",
        "word/comments.xml",
        "docProps/core.xml",
        "word/settings.xml",
        "word/_rels/document.xml.rels",
        "[Content_Types].xml",
        "_rels/.rels",
    ];

    /// <summary>Things that are easy to get wrong in a reader of WordprocessingML.</summary>
    internal static readonly string[] Tokens =
    [
        "<", ">", "\"", "/>", "</w:p>", "</w:r>", "</w:tc>", "</w:tr>", "</w:tbl>", "&amp;", "&#0;", "&#x1F;", "&#xD800;", "&lt;", "&bogus;",
        "<!DOCTYPE x [<!ENTITY e \"eeeeeeeeee\">]>", "&e;", "<![CDATA[x]]>", "<!-- c -->", "<?pi x?>",
        "<w:p/>", "<w:p><w:r><w:t>x</w:t></w:r></w:p>", "<w:r/>", "<w:r><w:t/></w:r>", "<w:t xml:space=\"preserve\"> x </w:t>", "<w:tab/>", "<w:br/>", "<w:br w:type=\"page\"/>", "<w:cr/>",
        "<w:sym w:char=\"F04A\"/>", "<w:sym w:char=\"FFFFFFFF\"/>", "<w:sym w:char=\"110000\"/>", "<w:sym w:char=\"D800\"/>", "<w:sym w:char=\"-1\"/>", "<w:sym/>",
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>", "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>", "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>",
        "<w:r><w:instrText>MERGEFIELD x</w:instrText></w:r>", "<w:r><w:instrText>HYPERLINK \"</w:instrText></w:r>", "<w:fldSimple w:instr=\"\"/>", "<w:fldSimple><w:r><w:t>f</w:t></w:r></w:fldSimple>",
        "<w:ins w:id=\"1\"><w:r><w:t>i</w:t></w:r></w:ins>", "<w:del w:id=\"2\"><w:r><w:delText>d</w:delText></w:r></w:del>", "<w:del>", "</w:del>", "<w:ins>", "</w:ins>",
        "<w:moveFrom><w:r><w:delText>m</w:delText></w:r></w:moveFrom>", "<w:moveTo><w:r><w:t>m</w:t></w:r></w:moveTo>",
        "<w:hyperlink r:id=\"rId4\"><w:r><w:t>l</w:t></w:r></w:hyperlink>", "<w:hyperlink r:id=\"rId999\">", "<w:hyperlink w:anchor=\"x\">", "<w:hyperlink r:id=\"rId5\">",
        "<w:sdt><w:sdtContent><w:p/></w:sdtContent></w:sdt>", "<w:sdt>", "<w:sdtContent>", "</w:sdtContent>",
        "<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>", "<w:tbl>", "<w:tr>", "<w:tc>", "<w:tcPr><w:gridSpan w:val=\"2147483647\"/></w:tcPr>", "<w:tcPr><w:gridSpan w:val=\"100000\"/></w:tcPr>",
        "<w:tcPr><w:gridSpan w:val=\"-5\"/></w:tcPr>", "<w:tcPr><w:vMerge/></w:tcPr>", "<w:tcPr><w:vMerge w:val=\"restart\"/></w:tcPr>", "<w:trPr><w:del w:id=\"3\"/></w:trPr>", "<w:trPr><w:ins w:id=\"4\"/></w:trPr>",
        "<w:pPr><w:numPr><w:ilvl w:val=\"99\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>", "<w:numPr><w:ilvl w:val=\"-1\"/><w:numId w:val=\"2\"/></w:numPr>", "<w:numId w:val=\"999\"/>",
        "<w:pStyle w:val=\"Heading1\"/>", "<w:pStyle w:val=\"\"/>", "<w:pStyle w:val=\"Loop\"/>", "<w:rStyle w:val=\"Hyperlink\"/>", "<w:outlineLvl w:val=\"99\"/>", "<w:outlineLvl w:val=\"-1\"/>",
        " w:val=\"2147483647\"", " w:val=\"-2147483648\"", " w:val=\"99999999999999999999\"", " w:val=\"1e309pt\"", " w:val=\"NaNpt\"", " w:val=\"\"", " w:val=\"0\"", " w:val=\"true\"", " w:val=\"off\"",
        "<w:sz w:val=\"99999999999999\"/>", "<w:ind w:left=\"99999999999999999\" w:hanging=\"-9223372036854775808\"/>", "<w:spacing w:before=\"9223372036854775807\" w:after=\"1e5cm\"/>",
        "<w:start w:val=\"2147483647\"/>", "<w:start w:val=\"-2147483648\"/>", "<w:start w:val=\"100000000\"/>", "<w:numFmt w:val=\"lowerLetter\"/>", "<w:numFmt w:val=\"upperRoman\"/>", "<w:numFmt w:val=\"bullet\"/>",
        "<w:lvlText w:val=\"%1.%2.%3.%4.%5.%6.%7.%8.%9.%0\"/>", "<w:lvlText w:val=\"%\"/>", "<w:lvl w:ilvl=\"8\">", "<w:lvl w:ilvl=\"9\">", "<w:lvl w:ilvl=\"-1\">", "<w:lvl>",
        "<w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"2147483647\"/></w:lvlOverride>", "<w:abstractNumId w:val=\"999\"/>", "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>",
        "<w:basedOn w:val=\"Normal\"/>", "<w:basedOn w:val=\"Heading1\"/>", "<w:style w:type=\"paragraph\" w:styleId=\"Normal\"><w:basedOn w:val=\"Normal\"/></w:style>", " w:styleId=\"Normal\"", " w:default=\"1\"",
        "<w:style w:type=\"character\" w:styleId=\"Loop\"><w:basedOn w:val=\"Loop\"/></w:style>",
        "<w:drawing><wp:inline><wp:extent cx=\"-1\" cy=\"9223372036854775807\"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed=\"rId999\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>",
        "<w:drawing><w:txbxContent><w:p><w:r><w:t>box</w:t></w:r></w:p></w:txbxContent></w:drawing>", "<w:txbxContent>", " r:embed=\"rId5\"", " r:embed=\"\"", " r:id=\"rId1\"", " r:id=\"\"",
        "<w:footnoteReference w:id=\"1\"/>", "<w:footnoteReference w:id=\"99\"/>", "<w:endnoteReference/>", "<w:commentReference w:id=\"0\"/>",
        "<w:headerReference w:type=\"default\" r:id=\"rId1\"/>", "<w:headerReference r:id=\"rId7\"/>", "<w:footerReference r:id=\"rId999\"/>", "<w:sectPr/>", "<w:sectPr><w:pgSz w:w=\"-1\" w:h=\"99999999999\"/></w:sectPr>",
        "<w:pgMar w:left=\"-99999999999999\" w:right=\"x\"/>", "<w:pgSz w:orient=\"landscape\"/>",
        "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"x\"><w:r><w:t>c</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>f</w:t></w:r></mc:Fallback></mc:AlternateContent>",
        " Target=\"../../../x.xml\"", " Target=\"/word/document.xml\"", " Target=\"\"", " Target=\"styles.xml\"", " TargetMode=\"External\"", " Target=\"header1.xml\"", " Target=\"c:\\windows\\win.ini\"", " Target=\"%2e%2e/x\"",
        " Id=\"rId1\"", " Id=\"\"", " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\"",
        " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"", " ContentType=\"\"", " PartName=\"/word/document.xml\"", " Extension=\"xml\"",
        "<dc:title>t</dc:title>", "<dcterms:created>9999-99-99</dcterms:created>", "<dcterms:modified>x</dcterms:modified>", " w:date=\"not a date\"", " w:date=\"0001-01-01T00:00:00Z\"",
        "<w:comment w:id=\"0\"><w:p/></w:comment>", "<w:footnote w:type=\"separator\" w:id=\"-1\"/>", "<w:footnote w:id=\"1\"><w:p><w:r><w:t>n</w:t></w:r></w:p></w:footnote>",
        "￿", "\uD800", "\0", "Öst", " xmlns=\"\"", " xmlns:w=\"http://purl.oclc.org/ooxml/wordprocessingml/main\"", " xmlns:w=\"urn:other\"",
        string.Concat(Enumerable.Repeat("<w:tbl><w:tr><w:tc>", 70)) + "<w:p/>" + string.Concat(Enumerable.Repeat("</w:tc></w:tr></w:tbl>", 70)),
        string.Concat(Enumerable.Repeat("<w:hyperlink>", 120)) + "<w:r><w:t>x</w:t></w:r>" + string.Concat(Enumerable.Repeat("</w:hyperlink>", 120)),
        string.Concat(Enumerable.Repeat("<x>", 5_000)) + string.Concat(Enumerable.Repeat("</x>", 5_000)),
        "<w:t>" + new string('a', 40_000) + "</w:t>",
    ];

    public static FuzzTarget Create()
    {
        return new FuzzTarget("docx-part", "one part of a valid document replaced; read with both readers, change, save, reopen", RunAsync, Generate);
    }

    private static byte[] Generate(Random random)
    {
        int part = random.Next(PartNames.Length);

        // The body, the styles and the numbering are where nearly all the parsing happens.
        if (random.Next(2) == 0)
        {
            part = random.Next(3);
        }

        string content = Template.Text(true, PartNames[part]);

        // One in ten unchanged: the save invariant matters most on a file that is not damaged,
        // and these are the best starting points for libFuzzer's corpus.
        if (random.Next(10) > 0)
        {
            content = Inputs.Mutate(random, content, Tokens);
        }

        byte[] text = Encoding.UTF8.GetBytes(content);
        byte[] input = new byte[text.Length + 2];
        input[0] = (byte)part;
        input[1] = (byte)random.Next(256);
        text.CopyTo(input, 2);

        return input;
    }

    internal static async Task RunAsync(byte[] input)
    {
        if (input.Length < 2)
        {
            return;
        }

        string part = PartNames[input[0] % PartNames.Length];
        await using MemoryStream file = Template.Build(Template.Rich, part, input.AsSpan(2));
        await Dump.WriteAsync("input.docx", file);
        await Harness.RunFileAsync(file.ToArray(), input[1]);
    }
}

/// <summary>Set by <c>--repro FILE --dump DIR</c>: the built file is written there, to open in Word and compare.</summary>
internal static class Dump
{
    public static string Directory { get; set; } = string.Empty;

    public static async Task WriteAsync(string name, MemoryStream file)
    {
        if (Directory.Length == 0)
        {
            return;
        }

        System.IO.Directory.CreateDirectory(Directory);
        await File.WriteAllBytesAsync(Path.Combine(Directory, name), file.ToArray());
        file.Position = 0;
    }
}
