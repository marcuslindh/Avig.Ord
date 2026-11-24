using System.IO.Compression;
using System.Text;

namespace Avig.Ord.Fuzz;

/// <summary>
/// The documents the targets start from. Every one of them is written by the package's own
/// <see cref="DocxWriter"/>, so that the seeds are files the package claims to understand. The
/// rich template then gets what the writer does not make: tracked changes, complex fields, a
/// symbol, content controls, a text box in alternate content, a comment and a comments part.
/// </summary>
internal static class Template
{
    public const string Wml = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    private const string Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

    /// <summary>A 2×1 PNG: the smallest picture with a size to read.</summary>
    public static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR4nGP4z8DwnwEKAA34A/1q6yb3AAAAAElFTkSuQmCC");

    public const string Svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"40\" height=\"20\"><rect width=\"40\" height=\"20\" fill=\"red\"/></svg>";

    /// <summary>What the rich template adds to the body, before its last section properties.</summary>
    public const string RichBody = $"""
        <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t xml:space="preserve">Det var en </w:t></w:r><w:del w:id="901" w:author="A" w:date="2026-11-24T00:00:00Z"><w:r><w:delText>gammal</w:delText></w:r></w:del><w:ins w:id="902" w:author="A" w:date="2026-11-24T00:00:00Z"><w:r><w:t>ny</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> text.</w:t></w:r></w:p>
        <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> MERGEFIELD Namn \* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>«Namn»</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:sym w:font="Wingdings" w:char="F04A"/><w:noBreakHyphen/><w:cr/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>IF </w:instrText></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>MERGEFIELD X</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>inre</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:instrText> = 1 "ja" "nej"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>ja</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
        <w:sdt><w:sdtPr><w:tag w:val="block"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>I kontroll</w:t></w:r></w:p></w:sdtContent></w:sdt>
        <w:p><w:commentRangeStart w:id="0"/><w:sdt><w:sdtPr><w:alias w:val="Namn"/><w:tag w:val="namn"/></w:sdtPr><w:sdtContent><w:r><w:t>Anna</w:t></w:r></w:sdtContent></w:sdt><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:r><mc:AlternateContent xmlns:mc="{Mc}"><mc:Choice xmlns:wps="{Wps}" Requires="wps"><w:drawing><wps:txbx><w:txbxContent><w:p><w:r><w:t>I rutan</w:t></w:r></w:p></w:txbxContent></wps:txbx></w:drawing></mc:Choice><mc:Fallback><w:pict><w:p><w:r><w:t>Reserv</w:t></w:r></w:p></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>
        <w:p><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr><w:smartTag w:uri="urn:x" w:element="x"><w:r><w:t>Djup punkt</w:t></w:r></w:smartTag><w:r><w:br w:type="column"/><w:t>efter</w:t></w:r></w:p>
        """;

    private const string Comments = $"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="{Wml}"><w:comment w:id="0" w:author="Eva" w:initials="E" w:date="2026-11-24T10:00:00Z"><w:p><w:r><w:t>Stämmer detta?</w:t></w:r></w:p></w:comment></w:comments>""";

    private static readonly Lazy<IReadOnlyList<(string Name, byte[] Data)>> PlainParts = new(() => ReadParts(Write(false)));

    private static readonly Lazy<IReadOnlyList<(string Name, byte[] Data)>> RichParts = new(() => MakeRich(PlainParts.Value));

    /// <summary>The parts of the plain template: exactly what <see cref="DocxWriter"/> wrote. It passes Microsoft's schema validator.</summary>
    public static IReadOnlyList<(string Name, byte[] Data)> Plain => PlainParts.Value;

    /// <summary>The parts of the rich template.</summary>
    public static IReadOnlyList<(string Name, byte[] Data)> Rich => RichParts.Value;

    /// <summary>The plain or the rich template.</summary>
    public static IReadOnlyList<(string Name, byte[] Data)> Parts(bool rich)
    {
        if (rich == true)
        {
            return Rich;
        }

        return Plain;
    }

    /// <summary>The text of a part of a template.</summary>
    public static string Text(bool rich, string name)
    {
        return Encoding.UTF8.GetString(Parts(rich).First(part => part.Name == name).Data);
    }

    /// <summary>A zip package of the template's parts, with one of them replaced (none when <paramref name="replaced"/> is empty).</summary>
    public static MemoryStream Build(IReadOnlyList<(string Name, byte[] Data)> parts, string replaced, ReadOnlySpan<byte> content)
    {
        MemoryStream stream = new();

        using (ZipArchive archive = new(stream, ZipArchiveMode.Create, true))
        {
            foreach ((string name, byte[] data) in parts)
            {
                ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Fastest);
                using Stream entryStream = entry.Open();

                if (name == replaced)
                {
                    entryStream.Write(content);
                }
                else
                {
                    entryStream.Write(data);
                }
            }
        }

        stream.Position = 0;

        return stream;
    }

    /// <summary>The template with its body replaced by <paramref name="body"/> (body content only; the last sectPr is kept).</summary>
    public static MemoryStream WithBody(bool rich, string body)
    {
        string document = Text(rich, "word/document.xml");
        int bodyStart = document.IndexOf("<w:body>", StringComparison.Ordinal) + "<w:body>".Length;
        int sectionStart = document.LastIndexOf("<w:sectPr", StringComparison.Ordinal);
        string replaced = string.Concat(document.AsSpan(0, bodyStart), body, document.AsSpan(sectionStart));

        return Build(Parts(rich), "word/document.xml", Encoding.UTF8.GetBytes(replaced));
    }

    /// <summary>The ids of the main part's relationships, by the last segment of their type: image, hyperlink, footnotes …</summary>
    public static IReadOnlyDictionary<string, string> MainRelationshipIds()
    {
        Dictionary<string, string> ids = new(StringComparer.Ordinal);
        string rels = Text(true, "word/_rels/document.xml.rels");
        int position = 0;

        while ((position = rels.IndexOf("<Relationship ", position, StringComparison.Ordinal)) >= 0)
        {
            int end = rels.IndexOf("/>", position, StringComparison.Ordinal);
            string element = rels[position..end];
            string id = Attribute(element, "Id");
            string type = Attribute(element, "Type");
            ids.TryAdd(type[(type.LastIndexOf('/') + 1)..], id);
            position = end;
        }

        return ids;
    }

    /// <summary>The every-feature document, as WriterTests.EveryFeatureAsync writes it, with a nested table, notes, fields and a template-like merge field.</summary>
    public static byte[] Write(bool flat)
    {
        using MemoryStream file = new();

        // The harness has no caller to hand a task to while the template is made the first time.
#pragma warning disable AVIG0014
        WriteAsync(file, flat).GetAwaiter().GetResult();
#pragma warning restore AVIG0014

        return file.ToArray();
    }

    private static async Task WriteAsync(MemoryStream file, bool flat)
    {
        await using DocxWriter writer = await DocxWriter.CreateAsync(file, new DocxWriterOptions { FlatOpc = flat, Title = "Rapporten", Author = "Test" }, leaveOpen: true);
        writer.Header(header => header.Paragraph("Sidhuvud"));
        writer.Header(header => header.Paragraph("Första sidan"), HeaderFooterKind.First);
        writer.Footer(footer => footer.Paragraph(paragraph => paragraph.Text("Sida ").PageNumber().Text(" av ").PageCount()));
        await writer.HeadingAsync("Årsrapport", 1);
        await writer.ParagraphAsync(paragraph => paragraph
            .Text("Vanlig ")
            .Text("fet", style => style.Bold().Size(14).Color("#c00000"))
            .Text(" kursiv", style => style.Italic().Underline().Strike())
            .Tab()
            .Text("efter tabb", style => style.Superscript().Font("Arial").Highlight("yellow"))
            .LineBreak()
            .Text("ny rad", style => style.AllCaps().SmallCaps()));
        await writer.ParagraphAsync(paragraph => paragraph.Link("Läs mer", "https://example.com/").Footnote("Källa: SCB.").Endnote("Slutnot.").Align(Alignment.Center).Spacing(Length.Pt(6), Length.Pt(12)).Indent(Length.Cm(1), Length.Cm(0.5), Length.Cm(-0.5)).KeepWithNext());
        await writer.ParagraphAsync(paragraph => paragraph.Text("Kära ").MergeField("Förnamn").Text(", ").Field("DATE \\@ \"yyyy-MM-dd\"", "2026-11-24"));
        await writer.BulletAsync("Ett");
        await writer.BulletAsync("Två");
        await writer.BulletAsync("Under", 1);
        await writer.NumberedAsync("Först");
        await writer.NumberedAsync("Sedan");
        await writer.NumberedAsync("Andra nivån", 1);
        await writer.NumberedAsync("Ny lista", startNew: true);
        await writer.TableAsync(table => table
            .HeaderRow("Namn", "Belopp", "Not")
            .Row(row => row.Cell("Hyra").Cell("8 000", cell => cell.Shading("D9D9D9")).Cell("Hela", cell => cell.MergeDown()))
            .Row(row => row.Cell("Summa", cell => cell.Span(2)).Cell(string.Empty, cell => cell.MergeUp()))
            .Row(row => row.Cell(cell => cell.Paragraph("Inre:").Table(inner => inner.Row("a", "b").Row("c", "d"))).Cell("x").Cell("y")));
        await writer.PictureAsync(Png, "En röd punkt");
        await writer.ParagraphAsync(paragraph => paragraph.Picture(Encoding.UTF8.GetBytes(Svg), Png, Length.Px(40), Length.Px(20), "svg"));
        await writer.SectionBreakAsync(PageSetup.A4.Turned());
        await writer.ParagraphAsync(paragraph => paragraph.Text("Liggande sida").PageBreak().Text("Nästa sida"));
        await writer.CompleteAsync();
    }

    private static List<(string Name, byte[] Data)> ReadParts(byte[] file)
    {
        List<(string Name, byte[] Data)> parts = [];
        using ZipArchive archive = new(new MemoryStream(file), ZipArchiveMode.Read);

        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            using Stream stream = entry.Open();
            using MemoryStream data = new();
            stream.CopyTo(data);
            parts.Add((entry.FullName, data.ToArray()));
        }

        return parts;
    }

    private static List<(string Name, byte[] Data)> MakeRich(IReadOnlyList<(string Name, byte[] Data)> plain)
    {
        List<(string Name, byte[] Data)> parts = [];

        foreach ((string name, byte[] data) in plain)
        {
            string text = Encoding.UTF8.GetString(data);

            if (name == "word/document.xml")
            {
                int section = text.LastIndexOf("<w:sectPr", StringComparison.Ordinal);
                text = text.Insert(section, RichBody.ReplaceLineEndings(string.Empty));
            }
            else if (name == "word/_rels/document.xml.rels")
            {
                text = text.Replace("</Relationships>", "<Relationship Id=\"rIdFuzzComments\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"comments.xml\"/></Relationships>", StringComparison.Ordinal);
            }
            else if (name == "[Content_Types].xml")
            {
                text = text.Replace("</Types>", "<Override PartName=\"/word/comments.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml\"/></Types>", StringComparison.Ordinal);
            }
            else
            {
                parts.Add((name, data));

                continue;
            }

            parts.Add((name, Encoding.UTF8.GetBytes(text)));
        }

        parts.Add(("word/comments.xml", Encoding.UTF8.GetBytes(Comments)));

        return parts;
    }

    private static string Attribute(string element, string name)
    {
        int start = element.IndexOf(" " + name + "=\"", StringComparison.Ordinal);

        if (start < 0)
        {
            return string.Empty;
        }

        start += name.Length + 3;

        return element[start..element.IndexOf('"', start)];
    }
}
