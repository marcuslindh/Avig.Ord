using System.IO.Compression;

namespace Avig.Ord.Tests;

public class WriterTests
{
    /// <summary>A 2×1 PNG: the smallest picture with a size to read.</summary>
    internal static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR4nGP4z8DwnwEKAA34A/1q6yb3AAAAAElFTkSuQmCC");

    internal const string Svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"40\" height=\"20\"><rect width=\"40\" height=\"20\" fill=\"red\"/></svg>";

    [Fact]
    public async Task A_file_with_every_feature_passes_Microsofts_schema_validation()
    {
        byte[] file = await EveryFeatureAsync(false);

        Assert.Equal(string.Empty, TestFiles.Validate(file));
    }

    [Fact]
    public async Task A_document_built_in_memory_with_every_feature_passes_Microsofts_schema_validation()
    {
        OrdDocument document = OrdDocument.Create();
        document.AddHeading("Rubrik", 1);
        document.AddParagraph(paragraph => paragraph.Text("Fet", style => style.Bold()).Link(" länk", "https://example.com/").Footnote("Not").Picture(Png, "bild"));
        document.AddParagraph(paragraph => paragraph.Bullet().Text("Punkt"));
        document.AddParagraph(paragraph => paragraph.Numbered().Text("Ett"));
        document.AddTable(table => table.HeaderRow("A", "B").Row("1", "2"));
        document.Properties.Title = "Titel";

        byte[] file = await TestFiles.SaveAsync(document);

        Assert.Equal(string.Empty, TestFiles.Validate(file));
    }

    [Fact]
    public async Task What_is_written_reads_back_as_written()
    {
        byte[] file = await EveryFeatureAsync(false);
        OrdDocument document = await TestFiles.OpenAsync(file);
        IReadOnlyList<Paragraph> paragraphs = document.Paragraphs;

        Paragraph heading = paragraphs[0];
        Assert.Equal("Årsrapport", heading.Text);
        Assert.Equal(0, heading.OutlineLevel);
        Assert.Equal("Heading1", heading.StyleId);

        Paragraph formatted = paragraphs[1];
        Assert.Equal("Vanlig fet kursiv\tefter tabb\nny rad", formatted.Text);
        Run bold = formatted.Runs.Single(run => run.Text == "fet");
        Assert.True(bold.Format.Bold);
        Assert.Equal(Length.Pt(14), bold.Format.Size);
        Assert.Equal("C00000", bold.Format.Color);
        Assert.True(formatted.Runs.Single(run => run.Text.StartsWith(" kursiv", StringComparison.Ordinal)).Format.Italic);

        Hyperlink link = Assert.Single(paragraphs[2].Inlines.OfType<Hyperlink>());
        Assert.Equal("https://example.com/", link.Url);
        Assert.Equal("Läs mer", link.Text);

        Assert.Equal(["• Ett", "• Två", "◦ Under", "1. Först", "2. Sedan", "1. Ny lista"], paragraphs.Where(paragraph => paragraph.List is not null).Select(paragraph => paragraph.List!.Label + " " + paragraph.Text));
        Assert.Equal(1, paragraphs.Single(paragraph => paragraph.Text == "Under").List!.Level);

        Table table = Assert.Single(document.Tables);
        Assert.Equal(3, table.ColumnCount);
        Assert.True(table.Rows[0].IsHeader);
        Assert.Equal("Namn\tBelopp\tNot", table.Rows[0].Text);
        Assert.Equal(2, table.Rows[2].Cells[0].ColumnSpan);
        Assert.Equal(VerticalMerge.Start, table.Rows[1].Cells[2].VerticalMerge);
        Assert.Equal(VerticalMerge.Continue, table.Rows[2].Cells[1].VerticalMerge);

        Picture picture = Assert.Single(document.AllPictures());
        Assert.Equal("En röd punkt", picture.Description);
        Assert.Equal(Length.Px(2), picture.Width);
        Assert.True((await picture.ReadAsync()).IsSuccess(out byte[]? png));
        Assert.Equal(Png, png);

        Note note = Assert.Single(document.Footnotes);
        Assert.Equal("Källa: SCB.", note.Text);
        Assert.Equal(note.Id, Assert.Single(paragraphs.SelectMany(paragraph => paragraph.Inlines).OfType<NoteReference>()).Id);

        Assert.Equal("Sidhuvud", Assert.Single(document.Headers).Text);
        Assert.Equal("Sida 1 av 1", Assert.Single(document.Footers).Text);
        Assert.Equal(["PAGE", "NUMPAGES"], document.Footers[0].Blocks.OfType<Paragraph>().SelectMany(paragraph => paragraph.Inlines).OfType<Field>().Select(field => field.Code));

        Assert.Equal(2, document.Sections.Count);
        Assert.Equal(Orientation.Portrait, document.Sections[0].Orientation);
        Assert.Equal(Orientation.Landscape, document.Sections[1].Orientation);
        Assert.Equal(Length.Cm(2.5).ToTwips(), document.Sections[0].MarginLeft!.Value.ToTwips());
        Assert.Equal("Rapporten", document.Properties.Title);
    }

    [Fact]
    public async Task The_same_picture_twice_is_stored_once()
    {
        using MemoryStream file = new();

        await using (DocxWriter writer = await DocxWriter.CreateAsync(file, leaveOpen: true))
        {
            await writer.PictureAsync(Png);
            await writer.PictureAsync(Png);
        }

        using ZipArchive archive = new(new MemoryStream(file.ToArray()));
        Assert.Single(archive.Entries, entry => entry.FullName.StartsWith("word/media/", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Flat_opc_reads_back_like_the_zip()
    {
        byte[] zip = await EveryFeatureAsync(false);
        byte[] flat = await EveryFeatureAsync(true);

        Assert.Equal('<', (char)flat[0]);
        Assert.Equal((await TestFiles.OpenAsync(zip)).ToPlainText(), (await TestFiles.OpenAsync(flat)).ToPlainText());
    }

    [Fact]
    public async Task An_svg_picture_has_its_png_to_fall_back_on()
    {
        using MemoryStream file = new();

        await using (DocxWriter writer = await DocxWriter.CreateAsync(file, leaveOpen: true))
        {
            await writer.ParagraphAsync(paragraph => paragraph.Picture(System.Text.Encoding.UTF8.GetBytes(Svg), Png, Length.Px(40), Length.Px(20), "svg"));
        }

        byte[] bytes = file.ToArray();
        Assert.Equal(string.Empty, TestFiles.Validate(bytes));
        OrdDocument document = await TestFiles.OpenAsync(bytes);
        Picture picture = Assert.Single(document.AllPictures());
        Assert.EndsWith(".png", picture.PartName, StringComparison.Ordinal);
        Assert.EndsWith(".svg", picture.SvgPartName, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Text_with_characters_xml_cannot_hold_is_written_without_them()
    {
        using MemoryStream file = new();

        await using (DocxWriter writer = await DocxWriter.CreateAsync(file, leaveOpen: true))
        {
            await writer.ParagraphAsync("a\u0001b\uD800c");
        }

        Assert.Equal("abc", (await TestFiles.OpenAsync(file.ToArray())).ToPlainText());
    }

    [Fact]
    public async Task A_template_is_written_as_a_template()
    {
        using MemoryStream file = new();

        await using (DocxWriter writer = await DocxWriter.CreateAsync(file, new DocxWriterOptions { Template = true }, leaveOpen: true))
        {
            await writer.ParagraphAsync("Mall");
        }

        Assert.Equal(WordFileKind.Template, (await TestFiles.OpenAsync(file.ToArray())).Kind);
    }

    [Fact]
    public void Lengths_convert_between_units_exactly()
    {
        Assert.Equal(25, Length.Cm(2.5).ToMm(), 9);
        Assert.Equal(72, Length.Inch(1).ToPt(), 9);
        Assert.Equal(1440, Length.Inch(1).ToTwips());
        Assert.Equal(22, Length.Pt(11).ToHalfPoints());
        Assert.Equal(96, Length.Inch(1).ToPx(), 9);
        Assert.True(Length.Cm(1) < Length.Inch(1));
        Assert.Equal("12pt", Length.Pt(12).ToString());
    }

    internal static async Task<byte[]> EveryFeatureAsync(bool flat)
    {
        using MemoryStream file = new();

        await using (DocxWriter writer = await DocxWriter.CreateAsync(file, new DocxWriterOptions { FlatOpc = flat, Title = "Rapporten", Author = "Test" }, leaveOpen: true))
        {
            writer.Header(header => header.Paragraph("Sidhuvud"));
            writer.Footer(footer => footer.Paragraph(paragraph => paragraph.Text("Sida ").PageNumber().Text(" av ").PageCount()));
            await writer.HeadingAsync("Årsrapport", 1);
            await writer.ParagraphAsync(paragraph => paragraph
                .Text("Vanlig ")
                .Text("fet", style => style.Bold().Size(14).Color("#c00000"))
                .Text(" kursiv", style => style.Italic())
                .Tab()
                .Text("efter tabb")
                .LineBreak()
                .Text("ny rad"));
            await writer.ParagraphAsync(paragraph => paragraph.Link("Läs mer", "https://example.com/").Footnote("Källa: SCB.").Align(Alignment.Center).Spacing(Length.Pt(6), Length.Pt(12)));
            await writer.BulletAsync("Ett");
            await writer.BulletAsync("Två");
            await writer.BulletAsync("Under", 1);
            await writer.NumberedAsync("Först");
            await writer.NumberedAsync("Sedan");
            await writer.NumberedAsync("Ny lista", startNew: true);
            await writer.TableAsync(table => table
                .HeaderRow("Namn", "Belopp", "Not")
                .Row(row => row.Cell("Hyra").Cell("8 000", cell => cell.Shading("D9D9D9")).Cell("Hela", cell => cell.MergeDown()))
                .Row(row => row.Cell("Summa", cell => cell.Span(2)).Cell(string.Empty, cell => cell.MergeUp())));
            await writer.PictureAsync(Png, "En röd punkt");
            await writer.SectionBreakAsync(PageSetup.A4.Turned());
            await writer.ParagraphAsync(paragraph => paragraph.Text("Liggande sida").PageBreak().Text("Nästa sida"));
            await writer.CompleteAsync();
        }

        return file.ToArray();
    }
}
