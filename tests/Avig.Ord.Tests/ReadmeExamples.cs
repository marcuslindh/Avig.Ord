using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace Avig.Ord.Tests;

#region example: typed-template-type
[OrdTemplate]
public partial class Offer
{
    public string Customer { get; init; } = "";

    [TemplateName("Valid until")]
    public DateOnly ValidUntil { get; init; }

    public IReadOnlyList<OfferLine> Lines { get; init; } = [];

    public decimal Total => Lines.Sum(line => line.Price);
}

[OrdTemplate]
public partial class OfferLine
{
    public string Item { get; init; } = "";

    public decimal Price { get; init; }
}
#endregion

/// <summary>
/// The code of every example in README.md. Each block there has a line
/// <c>&lt;!-- example: name --&gt;</c> before it and must be exactly the code of
/// <c>#region example: name</c> here; <see cref="Every_README_example_is_code_that_runs_here"/>
/// checks that. The examples use relative paths, so each test runs in a folder of its own, and
/// the collection runs alone, since the current directory belongs to the whole process.
/// </summary>
[Collection(ReadmeGroup.Name)]
public sealed class ReadmeExamples : IDisposable
{
    private readonly string previousDirectory = Environment.CurrentDirectory;
    private readonly string folder = Path.Combine(Path.GetTempPath(), "avig-ord-readme-" + Guid.NewGuid().ToString("N"));

    public ReadmeExamples()
    {
        Directory.CreateDirectory(folder);
        Environment.CurrentDirectory = folder;
    }

    public void Dispose()
    {
        Environment.CurrentDirectory = previousDirectory;
        Directory.Delete(folder, true);
    }

    [Fact]
    public async Task Write_a_document()
    {
        byte[] logo = WriterTests.Png;

        #region example: write
        await using (DocxWriter writer = await DocxWriter.CreateAsync(File.Create("report.docx")))
        {
            writer.Header(header => header.Paragraph("Quarterly report"));
            writer.Footer(footer => footer.Paragraph(paragraph => paragraph.Text("Page ").PageNumber().Text(" of ").PageCount()));

            await writer.HeadingAsync("Sales", 1);
            await writer.ParagraphAsync(paragraph => paragraph
                .Text("Sales rose by ")
                .Text("12 %", style => style.Bold().Color("2E7D32"))
                .Text(" on last quarter.")
                .Footnote("Source: the sales system, 2026-10-01."));

            await writer.BulletAsync("North: 4 200");
            await writer.BulletAsync("South: 3 900");

            await writer.TableAsync(table => table
                .HeaderRow("Region", "Q2", "Q3")
                .Row("North", "3 700", "4 200")
                .Row("South", "3 600", "3 900"));

            await writer.PictureAsync(logo, "Company logo");
            await writer.SectionBreakAsync(PageSetup.A4.Turned());   // landscape pages from here
            await writer.ParagraphAsync("Appendix");
        }
        #endregion

        OrdDocument document = await OpenAsync("report.docx");
        Assert.Equal("Sales", document.Paragraphs[0].Text);
        TestFiles.AssertValid(await File.ReadAllBytesAsync("report.docx"));
    }

    [Fact]
    public async Task Read_a_document()
    {
        await File.WriteAllBytesAsync("report.docx", await WriterTests.EveryFeatureAsync(false));

        #region example: read
        Result<OrdDocument> opened = await OrdDocument.OpenAsync("report.docx");

        if (opened.IsFailure(out string? error, out _, out OrdDocument? document) == true)
        {
            Console.WriteLine(error);

            return;
        }

        foreach (Paragraph paragraph in document.Paragraphs)
        {
            string indent = new(' ', (paragraph.List?.Level ?? 0) * 2);

            if (paragraph.IsHeading == true)
            {
                Console.WriteLine($"Heading {paragraph.OutlineLevel + 1}: {paragraph.Text}");
            }
            else if (paragraph.List is ListInfo list)
            {
                Console.WriteLine($"{indent}{list.Label} {paragraph.Text}");   // "1. First", "a) Sub", "• Item"
            }
        }

        foreach (Table table in document.Tables)
        {
            foreach (TableRow row in table.Rows)
            {
                Console.WriteLine(string.Join(" | ", row.Cells.Select(cell => cell.Text)));
            }
        }

        // The formatting a run has as Word shows it: defaults, style, character style and its own.
        Run first = document.Paragraphs[1].Runs[0];
        Console.WriteLine($"{first.Text}: bold {first.Format.Bold}, {first.Format.Size?.ToPt()} pt");
        #endregion

        Assert.Equal(3, document.Paragraphs.Count(paragraph => paragraph.List?.IsBullet == true));
    }

    [Fact]
    public async Task Text_and_markdown()
    {
        await File.WriteAllBytesAsync("report.docx", await WriterTests.EveryFeatureAsync(false));

        #region example: text-and-markdown
        if ((await OrdDocument.OpenAsync("report.docx")).IsSuccess(out OrdDocument? document) == false)
        {
            return;
        }

        string text = document.ToPlainText();          // paragraphs on lines, cells separated by tabs
        string markdown = document.ToMarkdown();       // # headings, - lists, | tables |, **bold**, ![pictures](media/image1.png)

        // The Markdown file, and its pictures in media/ beside it.
        await document.SaveMarkdownAsync("report.md");
        #endregion

        Assert.Contains("Årsrapport", text, StringComparison.Ordinal);
        Assert.StartsWith("# Årsrapport", markdown, StringComparison.Ordinal);
        Assert.True(File.Exists(Path.Combine("media", "image1.png")));
    }

    [Fact]
    public async Task Stream_a_large_document()
    {
        await File.WriteAllBytesAsync("large.docx", await WriterTests.EveryFeatureAsync(false));

        #region example: stream
        // One paragraph or table at a time; the rest of the file is never in memory.
        Result<DocxReader> opened = await DocxReader.OpenAsync("large.docx");

        if (opened.IsSuccess(out DocxReader? reader) == true)
        {
            await using (reader)
            {
                await foreach (Block block in reader.ReadBlocksAsync())
                {
                    if (block is Paragraph { IsHeading: true } heading)
                    {
                        Console.WriteLine(heading.Text);
                    }
                }
            }
        }
        #endregion

        Assert.NotNull(reader);
    }

    [Fact]
    public async Task Edit_a_document()
    {
        await File.WriteAllBytesAsync("contract.docx", TestFiles.Docx("<w:p><w:r><w:t>Signed in 2025 by [NAME].</w:t></w:r></w:p><w:p><w:r><w:t>[PRICES]</w:t></w:r></w:p><w:sectPr/>"));

        #region example: edit
        if ((await OrdDocument.OpenAsync("contract.docx")).IsSuccess(out OrdDocument? document) == false)
        {
            return;
        }


        // Changes are saved as tracked changes, which can be read, accepted or rejected in Word.
        document.TrackChanges("Contract robot");

        // Found and replaced also where Word has split the text over several runs.
        document.Replace("2025", "2026");
        document.Replace("[NAME]", Replacement.Text("Anna Svensson", style => style.Bold()));
        document.Replace(new Regex(@"\[PRICES\]"), _ => Replacement.Table(table => table
            .HeaderRow("Service", "Price")
            .Row("Support", "1 200 kr")));

        document.Paragraphs[0].InsertParagraphAfter("Appendix 1 is part of this contract.");

        await document.SaveAsync("contract-v2.docx");
        #endregion

        OrdDocument saved = await OpenAsync("contract-v2.docx");
        Assert.StartsWith("Signed in 2026 by Anna Svensson.", saved.ToPlainText(), StringComparison.Ordinal);
        TestFiles.AssertValid(await File.ReadAllBytesAsync("contract-v2.docx"));
    }

    [Fact]
    public async Task Build_in_memory()
    {
        #region example: build
        OrdDocument document = OrdDocument.Create();
        document.Properties.Title = "Minutes";

        document.AddHeading("Minutes 24 November", 1);
        Paragraph present = document.AddParagraph("Present: Anna, Bo");
        present.AppendText(" (chair: Anna)", style => style.Italic());

        Table decisions = document.AddTable(table => table.HeaderRow("Item", "Decision").Row("1", "Approved"));
        decisions.AddRow("2", "Postponed");

        await document.SaveAsync("minutes.docx");
        #endregion

        TestFiles.AssertValid(await File.ReadAllBytesAsync("minutes.docx"));
    }

    [Fact]
    public async Task Fill_a_template()
    {
        await File.WriteAllBytesAsync("letter-template.docx", TestFiles.Docx(
            """
            <w:p><w:r><w:t>Dear {{Name}},</w:t></w:r></w:p>
            <w:p><w:r><w:t>your balance on {{Date:d MMMM yyyy}} is {{Balance:N2}} kr.</w:t></w:r></w:p>
            <w:p><w:r><w:t>{{#Overdue}}</w:t></w:r></w:p><w:p><w:r><w:t>Please pay within ten days.</w:t></w:r></w:p><w:p><w:r><w:t>{{/Overdue}}</w:t></w:r></w:p>
            """));

        #region example: template
        if ((await OrdDocument.OpenAsync("letter-template.docx")).IsSuccess(out OrdDocument? letter) == false)
        {
            return;
        }

        TemplateValues values = new TemplateValues()
            .Set("Name", "Anna")
            .Set("Date", new DateOnly(2026, 11, 24))
            .Set("Balance", 1200.5m)
            .Set("Overdue", true);

        TemplateResult result = letter.FillTemplate(values, new TemplateOptions { Culture = CultureInfo.GetCultureInfo("sv-SE") });

        // A name without a value becomes empty; it is not an error, but it is listed.
        foreach (MissingValue missing in result.Missing)
        {
            Console.WriteLine($"No value for {missing.Name} in {missing.Location}");
        }

        await letter.SaveAsync("letter-anna.docx");
        #endregion

        OrdDocument saved = await OpenAsync("letter-anna.docx");
        Assert.Equal("Dear Anna,\nyour balance on 24 november 2026 is 1 200,50 kr.\nPlease pay within ten days.", saved.ToPlainText());
    }

    [Fact]
    public async Task Fill_a_typed_template()
    {
        await File.WriteAllBytesAsync("offer-template.docx", TestFiles.Docx(
            """
            <w:p><w:r><w:t>Offer for {{Customer}}, valid until {{Valid until:yyyy-MM-dd}}</w:t></w:r></w:p>
            <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
              <w:tr><w:tc><w:p><w:r><w:t>{{#Lines}}{{Item}}</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>{{Price:0.00}}{{/Lines}}</w:t></w:r></w:p></w:tc></w:tr>
            </w:tbl>
            <w:p><w:r><w:t>Total {{Total:0.00}}</w:t></w:r></w:p>
            """));

        #region example: typed-template
        Offer offer = new()
        {
            Customer = "Anna AB",
            ValidUntil = new DateOnly(2026, 12, 31),
            Lines = [new() { Item = "Installation", Price = 4500 }, new() { Item = "Support", Price = 1200 }],
        };

        if ((await OrdDocument.OpenAsync("offer-template.docx")).IsSuccess(out OrdDocument? document) == false)
        {
            return;
        }

        document.FillTemplate(offer, new TemplateOptions { Culture = CultureInfo.InvariantCulture });   // one table row per line
        await document.SaveAsync("offer.docx");
        #endregion

        OrdDocument saved = await OpenAsync("offer.docx");
        Assert.Equal("Offer for Anna AB, valid until 2026-12-31\nInstallation\t4500.00\nSupport\t1200.00\nTotal 5700.00", saved.ToPlainText());
    }

    [Fact]
    public async Task Tracked_changes()
    {
        await File.WriteAllBytesAsync("reviewed.docx", TestFiles.Docx(ReaderTests.TrackedBody));

        #region example: revisions
        // Accepted (the default) reads the text as if every change were accepted, Original as it
        // was before them, and Marked has both, each run marked.
        ReadOptions marked = new() { Revisions = RevisionView.Marked };

        if ((await OrdDocument.OpenAsync("reviewed.docx", marked)).IsSuccess(out OrdDocument? document) == false)
        {
            return;
        }

        foreach (Run run in document.Paragraphs[0].Runs.Where(run => run.Revision != RevisionKind.None))
        {
            Console.WriteLine($"{run.Revision}: {run.Text}");   // Deleted: old, Inserted: new
        }

        document.AcceptAllRevisions();   // or RejectAllRevisions(), before saving a clean copy
        #endregion

        Assert.Equal("Det var en ny text.", document.ToPlainText());
    }

    [Fact]
    public async Task Clean_a_file_from_outside()
    {
        await File.WriteAllBytesAsync("from-outside.docm", await ActiveContentTests.FileWithAsync("everything"));

        #region example: clean
        if ((await OrdDocument.OpenAsync("from-outside.docm")).IsSuccess(out OrdDocument? document) == false)
        {
            return;
        }

        Console.WriteLine(document.ActiveContent);   // Macros, ActiveX, EmbeddedObjects, ActiveFields, ExternalReferences

        foreach (Avig.OpenXml.OpcExternalRelationship link in document.ExternalReferences)
        {
            Console.WriteLine($"{link.Kind}: {link.Target}");   // attachedTemplate: https://evil.example/template.dotm
        }

        // Macros, ActiveX controls, embedded objects, fields such as DDE and every external
        // reference but hyperlinks go; the text and the pictures stay.
        document.RemoveActiveContent();
        await document.SaveAsync("cleaned.docx");
        #endregion

        OrdDocument cleaned = await OpenAsync("cleaned.docx");
        Assert.Equal(ActiveContent.None, cleaned.ActiveContent);
        TestFiles.AssertValid(await File.ReadAllBytesAsync("cleaned.docx"));
    }

    private static async Task<OrdDocument> OpenAsync(string path)
    {
        Result<OrdDocument> opened = await OrdDocument.OpenAsync(path);
        Assert.True(opened.IsSuccess(out OrdDocument? document), opened.Error);

        return document;
    }

    [Fact]
    public async Task Upload()
    {
        string answer = await HandleUploadAsync(new MemoryStream("not a document"u8.ToArray()));

        Assert.Equal("Upload a .docx file.", answer);
    }

    [Fact]
    public void Every_README_example_is_code_that_runs_here()
    {
        string source = File.ReadAllText(SourcePath());
        string readme = File.ReadAllText(Path.Combine(Path.GetDirectoryName(SourcePath())!, "..", "..", "README.md"));
        Dictionary<string, string> regions = [];

        foreach (Match match in Regex.Matches(source, @"^[ \t]*#region example: (?<name>[\w-]+)\r?\n(?<code>.*?)^[ \t]*#endregion", RegexOptions.Multiline | RegexOptions.Singleline))
        {
            regions[match.Groups["name"].Value] = Dedent(match.Groups["code"].Value);
        }

        Dictionary<string, string> blocks = [];

        foreach (Match match in Regex.Matches(readme, @"<!-- example: (?<name>[\w-]+) -->\r?\n```csharp\r?\n(?<code>.*?)```", RegexOptions.Singleline))
        {
            blocks[match.Groups["name"].Value] = Normalize(match.Groups["code"].Value);
        }

        Assert.Equal(regions.Keys.Order(StringComparer.Ordinal), blocks.Keys.Order(StringComparer.Ordinal));

        foreach ((string name, string code) in regions)
        {
            Assert.True(code == blocks[name], $"The README example '{name}' is not the code in ReadmeExamples.cs.\n--- README ---\n{blocks[name]}\n--- test ---\n{code}");
        }

        // Every csharp block in the README is a tested example.
        Assert.Equal(blocks.Count, Regex.Count(readme, "```csharp"));
    }

    private static async Task<string> HandleUploadAsync(Stream upload)
    {
        #region example: upload
        // Lower limits for files from outside, and a time limit.
        OrdLimits limits = OrdLimits.Default with { MaxUncompressedBytes = 50_000_000 };

        using CancellationTokenSource timeout = new(TimeSpan.FromSeconds(10));

        Result<OrdDocument> opened = await OrdDocument.OpenAsync(upload, new ReadOptions { Limits = limits }, timeout.Token);

        if (opened.IsFailure(out string? error, out _, out OrdDocument? document) == true)
        {
            // The error names parts and quotes text: log it, and show the user a message of your own.
            Console.WriteLine($"Upload rejected: {error}");

            return opened.ErrorKind() switch
            {
                OrdErrorKind.TooLarge => "The file is too large.",
                OrdErrorKind.NotWord => "Upload a .docx file.",
                OrdErrorKind.Damaged => "The file is damaged.",
                _ => "The file could not be read.",
            };
        }

        return $"{document.Paragraphs.Count} paragraphs";
        #endregion
    }

    private static string SourcePath([CallerFilePath] string path = "")
    {
        return path;
    }

    private static string Dedent(string code)
    {
        string[] lines = Normalize(code).Split('\n');
        int indent = lines.Where(line => line.Trim().Length > 0).Select(line => line.Length - line.TrimStart().Length).DefaultIfEmpty(0).Min();
        List<string> dedented = [];

        foreach (string line in lines)
        {
            if (line.Length >= indent)
            {
                dedented.Add(line[indent..]);
            }
            else
            {
                dedented.Add(line.TrimStart());
            }
        }

        return Normalize(string.Join('\n', dedented));
    }

    private static string Normalize(string code)
    {
        return code.Replace("\r\n", "\n", StringComparison.Ordinal).TrimEnd() + "\n";
    }
}

/// <summary>The README examples change the current directory, so they run alone.</summary>
[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class ReadmeGroup
{
    public const string Name = "README examples";
}
