using System.Text;

namespace Avig.Ord.WordCheck;

/// <summary>The documents Avig.Ord writes for Word to open: with each of its ways to write a file.</summary>
internal static class OrdWritten
{
    /// <summary>A 2×1 PNG.</summary>
    public static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR4nGP4z8DwnwEKAA34A/1q6yb3AAAAAElFTkSuQmCC");

    public const string Svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"40\" height=\"20\"><rect width=\"40\" height=\"20\" fill=\"red\"/></svg>";

    /// <summary>Writes every document into <paramref name="folder"/> and gives their paths. The Word files in <paramref name="wordFiles"/> are opened, changed and saved again too.</summary>
    public static async Task<List<string>> WriteAllAsync(string folder, string wordFiles)
    {
        Directory.CreateDirectory(folder);
        List<string> paths = [];

        string everyFeature = Path.Combine(folder, "writer-every-feature.docx");
        await File.WriteAllBytesAsync(everyFeature, await EveryFeatureAsync(new DocxWriterOptions { Title = "Rapporten", Author = "Avig.Ord" }));
        paths.Add(everyFeature);

        string flat = Path.Combine(folder, "writer-flat-opc.xml");
        await File.WriteAllBytesAsync(flat, await EveryFeatureAsync(new DocxWriterOptions { FlatOpc = true }));
        paths.Add(flat);

        string template = Path.Combine(folder, "writer-template.dotx");
        await File.WriteAllBytesAsync(template, await EveryFeatureAsync(new DocxWriterOptions { Template = true }));
        paths.Add(template);

        string created = Path.Combine(folder, "model-created.docx");
        await Check(await CreatedDocument().SaveAsync(created));
        paths.Add(created);

        string edited = Path.Combine(folder, "model-opened-edited.docx");
        OrdDocument opened = await OrdSide.OpenAsync(everyFeature, RevisionView.Accepted);
        Edit(opened);
        await Check(await opened.SaveAsync(edited));
        paths.Add(edited);

        string editedFlat = Path.Combine(folder, "model-opened-edited-flat-opc.xml");
        await Check(await opened.SaveAsync(editedFlat, new SaveOptions { FlatOpc = true }));
        paths.Add(editedFlat);

        if (Directory.Exists(wordFiles) == true)
        {
            foreach (string wordFile in Directory.GetFiles(wordFiles, "*.docx").Order(StringComparer.Ordinal))
            {
                string saved = Path.Combine(folder, "roundtrip-" + Path.GetFileName(wordFile));
                OrdDocument document = await OrdSide.OpenAsync(wordFile, RevisionView.Marked);
                document.AddParagraph("Tillagd av Avig.Ord.");
                await Check(await document.SaveAsync(saved));
                paths.Add(saved);
            }
        }

        return paths;
    }

    private static async Task Check(Result saved)
    {
        await Task.CompletedTask;

        if (saved.IsFailure(out string? error, out _) == true)
        {
            throw new InvalidOperationException(error);
        }
    }

    /// <summary>As WriterTests.EveryFeatureAsync, with endnotes, fields, a nested table and an SVG picture too.</summary>
    private static async Task<byte[]> EveryFeatureAsync(DocxWriterOptions options)
    {
        using MemoryStream file = new();

        await using (DocxWriter writer = await DocxWriter.CreateAsync(file, options, leaveOpen: true))
        {
            writer.Header(header => header.Paragraph("Sidhuvud"));
            writer.Footer(footer => footer.Paragraph(paragraph => paragraph.Text("Sida ").PageNumber().Text(" av ").PageCount()));
            await writer.HeadingAsync("Årsrapport", 1);
            await writer.ParagraphAsync(paragraph => paragraph
                .Text("Vanlig ")
                .Text("fet", style => style.Bold().Size(14).Color("#c00000"))
                .Text(" kursiv", style => style.Italic().Underline())
                .Tab()
                .Text("efter tabb", style => style.Strike().Highlight("yellow"))
                .LineBreak()
                .Text("ny rad", style => style.SmallCaps()));
            await writer.ParagraphAsync(paragraph => paragraph.Link("Läs mer", "https://example.com/").Footnote("Källa: SCB.").Endnote("Slutnot.").Align(Alignment.Center).Spacing(Length.Pt(6), Length.Pt(12)));
            await writer.ParagraphAsync(paragraph => paragraph.Text("Kära ").MergeField("Förnamn").Text("!"));
            await writer.HeadingAsync("Listor", 2);
            await writer.BulletAsync("Ett");
            await writer.BulletAsync("Två");
            await writer.BulletAsync("Under", 1);
            await writer.NumberedAsync("Först");
            await writer.NumberedAsync("Sedan");
            await writer.NumberedAsync("Andra nivån", 1);
            await writer.NumberedAsync("Tillbaka");
            await writer.NumberedAsync("Ny lista", startNew: true);
            await writer.TableAsync(table => table
                .HeaderRow("Namn", "Belopp", "Not")
                .Row(row => row.Cell("Hyra").Cell("8 000", cell => cell.Shading("D9D9D9")).Cell("Hela", cell => cell.MergeDown()))
                .Row(row => row.Cell("Summa", cell => cell.Span(2)).Cell(string.Empty, cell => cell.MergeUp()))
                .Row(row => row.Cell(cell => cell.Paragraph("Inre tabell:").Table(inner => inner.Row("a", "b").Row("c", "d"))).Cell("x").Cell("y")));
            await writer.PictureAsync(Png, "En röd punkt");
            await writer.ParagraphAsync(paragraph => paragraph.Picture(Encoding.UTF8.GetBytes(Svg), Png, Length.Px(40), Length.Px(20), "svg"));
            await writer.SectionBreakAsync(PageSetup.A4.Turned());
            await writer.ParagraphAsync(paragraph => paragraph.Text("Liggande sida").PageBreak().Text("Nästa sida"));
            await writer.CompleteAsync();
        }

        return file.ToArray();
    }

    private static OrdDocument CreatedDocument()
    {
        OrdDocument document = OrdDocument.Create();
        document.AddHeading("Protokoll", 1);
        document.AddParagraph(paragraph => paragraph.Text("Fet", style => style.Bold()).Link(" länk", "https://example.com/").Footnote("En not.").Picture(Png, "bild"));
        document.AddParagraph(paragraph => paragraph.Bullet().Text("Punkt"));
        document.AddParagraph(paragraph => paragraph.Bullet(1).Text("Underpunkt"));
        document.AddParagraph(paragraph => paragraph.Numbered().Text("Ett"));
        document.AddParagraph(paragraph => paragraph.Numbered().Text("Två"));
        Table table = document.AddTable(builder => builder.HeaderRow("A", "B").Row("1", "2"));
        table.AddRow("3", "4");
        table.InsertParagraphAfter("Efter tabellen");
        Paragraph removed = document.AddParagraph("Tas bort");
        removed.Remove();
        document.Paragraphs[0].InsertParagraphAfter("Inledning");
        document.Properties.Title = "Protokoll";

        return document;
    }

    private static void Edit(OrdDocument document)
    {
        IReadOnlyList<Paragraph> paragraphs = document.Paragraphs;
        paragraphs[0].AppendText(" 2026", style => style.Italic());
        paragraphs[1].InsertParagraphBefore("Ny första rad");
        paragraphs[2].SetText("Ersatt text");
        document.Tables[0].AddRow("Ny", "rad", "här");
        document.Tables[0].Rows[1].Cells[0].SetText("Lokal");
        document.AddParagraph(paragraph => paragraph.Numbered().Text("Ny punkt i ny lista"));
        document.AddTable(table => table.HeaderRow("C", "D").Row("5", "6"));
        document.AddParagraph("Slut.");
        document.Properties.Title = "Redigerad";
    }
}
