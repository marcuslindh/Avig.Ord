# Avig.Ord

Read, write and edit Word documents (.docx, .dotx, .docm and Flat OPC) in .NET, without Word and
without other packages than [Avig.Result](https://www.nuget.org/packages/Avig.Result) and
[Avig.OpenXml](https://www.nuget.org/packages/Avig.OpenXml), the package layer it shares with
[Avig.Kalkyl](https://www.nuget.org/packages/Avig.Kalkyl) for Excel. Trimming and Native AOT safe.

- **Read** paragraphs, runs, tables, lists with their labels, headings, pictures, hyperlinks,
  fields, footnotes, comments, headers and footers, content controls and text boxes, with the
  formatting a run really has: the document defaults, the style chain and its own.
- **Write** a document block by block with constant memory, or **build** one in memory.
- **Edit** a document and keep what the package does not understand (SmartArt, charts,
  equations, other programs' parts) as it was.
- **Search and replace** also where Word has split the text over runs, with regex, and replace
  with text, a picture, paragraphs or a table.
- **Track changes**: read them as accepted, original or marked; accept or reject them; make your
  own edits as tracked changes.
- **Templates**: `{{Name}}` placeholders with formats, `{{#Lines}}…{{/Lines}}` sections that
  repeat table rows or paragraphs, conditions, MERGEFIELD fields and content controls, filled from
  a dictionary or from a class through a source generator.
- **Text out**: plain text, or GitHub Flavored Markdown with headings, lists, tables and pictures.
- **Files from outside** are opened with limits on size, parts and XML depth, without DTDs, and
  every failure has a kind you can show a user.

What it does not do: the old binary .doc format, PDF or page layout, and updating fields and
tables of contents (fields keep the result Word saved).

## Write a document

`DocxWriter` writes straight to the file: only the block being written, and the pictures, are in
memory.

<!-- example: write -->
```csharp
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
```

## Read a document

Paragraphs and tables are thin views over the document: `Text`, `Runs`, `Format` (as Word shows
it), `DirectFormat` (what the run itself sets), `StyleName`, `OutlineLevel`, `List` with the label
Word shows, and for a table `Rows`, `Cells`, `ColumnSpan` and `VerticalMerge`. Every index is
0-based, and every length is a `Length` (`Length.Cm(2.5)`, `Length.Pt(11)`, `.ToCm()`), so the
unit is never a guess.

<!-- example: read -->
```csharp
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
```

Every failure you should expect is a `Result`, not an exception: a missing file is `NotFound`, a
file another program holds is `Conflict`, and a file that is not a Word document, is damaged or is
too large is `Invalid`. `ErrorKind()` says which, in words you can show a user:

<!-- example: upload -->
```csharp
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
```

## Text and Markdown

A paragraph with an outline level (Heading 1, Rubrik 1 or a style of your own) is a heading; lists
come out as `-` and `1.`, tables as GFM tables with the first row as header, emphasis as `**` and
`*`, and pictures as links into `media/`.

<!-- example: text-and-markdown -->
```csharp
if ((await OrdDocument.OpenAsync("report.docx")).IsSuccess(out OrdDocument? document) == false)
{
    return;
}

string text = document.ToPlainText();          // paragraphs on lines, cells separated by tabs
string markdown = document.ToMarkdown();       // # headings, - lists, | tables |, **bold**, ![pictures](media/image1.png)

// The Markdown file, and its pictures in media/ beside it.
await document.SaveMarkdownAsync("report.md");
```

## A large document

`DocxReader` reads one paragraph or table at a time, as read-only objects of the same types.
`ReadEventsAsync` goes further, element by element, for tables too large to hold at once.

<!-- example: stream -->
```csharp
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
```

## Edit a document

<!-- example: edit -->
```csharp
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
```

## Build a document in memory

<!-- example: build -->
```csharp
OrdDocument document = OrdDocument.Create();
document.Properties.Title = "Minutes";

document.AddHeading("Minutes 24 November", 1);
Paragraph present = document.AddParagraph("Present: Anna, Bo");
present.AppendText(" (chair: Anna)", style => style.Italic());

Table decisions = document.AddTable(table => table.HeaderRow("Item", "Decision").Row("1", "Approved"));
decisions.AddRow("2", "Postponed");

await document.SaveAsync("minutes.docx");
```

## Templates

A template is an ordinary Word document with `{{Name}}` where values go. `{{Name:format}}`
formats a number or date in the culture you give (the machine's when you give none), and
`{{Customer.Name}}` reaches into a nested value. `{{#Name}}` … `{{/Name}}` repeats what is
between for each item of a list — table rows when the tags are in a row, whole paragraphs when
they stand on paragraphs of their own — or keeps it once for `true`; `{{^Name}}` … `{{/Name}}`
keeps it when the value is false, empty or missing. Word's own MERGEFIELD fields (with `\@`,
`\#`, `\*`, `\b` and `\f`) and content controls (by tag, or title) are filled by name too, and a
`Replacement` as a value puts in a picture or a table.

<!-- example: template -->
```csharp
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
```

With a class instead of a dictionary, the source generator makes the mapping at compile time:

<!-- example: typed-template-type -->
```csharp
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
```

<!-- example: typed-template -->
```csharp
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
```

## Tracked changes

<!-- example: revisions -->
```csharp
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
```

## Files from outside

Avig.Ord never runs a macro, follows a link or updates a field, but a document it saves keeps
whatever it had. `ActiveContent` says what in a file can act on its own or fetch something when
Word opens it — the ways malware gets into documents: macros, ActiveX controls, embedded OLE
objects, fields such as DDE and INCLUDEPICTURE, and external references such as a template
attached from a server. `RemoveActiveContent` takes all of it out, and Word opens the result
without repair. Zip files that say one thing to Office and another to a scanner are refused when
they are opened.

<!-- example: clean -->
```csharp
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
```

## The API at a glance

| Task | API |
| --- | --- |
| Open, save | `OrdDocument.OpenAsync`, `SaveAsync` (with `SaveOptions.Format` and `FlatOpc`), `OrdDocument.Create` |
| Read in a stream | `DocxReader.OpenAsync`, `ReadBlocksAsync`, `ReadEventsAsync`, `ReadPlainTextAsync`, `WriteMarkdownAsync` |
| Write in a stream | `DocxWriter.CreateAsync`, `ParagraphAsync`, `HeadingAsync`, `BulletAsync`, `NumberedAsync`, `TableAsync`, `PictureAsync`, `PageBreakAsync`, `SectionBreakAsync`, `Header`, `Footer` |
| Content | `Blocks`, `Paragraphs`, `Tables`, `Sections`, `Headers`, `Footers`, `Footnotes`, `Endnotes`, `Comments`, `Properties`, `Styles` |
| A paragraph | `Text`, `Inlines`, `Runs`, `StyleId`, `StyleName`, `Format`, `DirectFormat`, `OutlineLevel`, `List` |
| Inlines | `Run`, `Hyperlink`, `Field`, `Picture` (`ReadAsync`), `Break`, `NoteReference`, `TextBox`, `ContentControl` |
| Change | `AddParagraph`, `AddHeading`, `AddTable`, `InsertParagraphAfter`/`Before`, `InsertTableAfter`, `AppendText`, `SetText`, `SetStyle`, `Remove`, `AddRow`, `InsertCopyAfter` |
| Search | `Find`, `Replace` (text, regex, `Replacement`), `SearchOptions` (case, whole word, scope) |
| Changes | `TrackChanges`, `StopTrackingChanges`, `AcceptAllRevisions`, `RejectAllRevisions`, `ReadOptions.Revisions` |
| Templates | `FillTemplate`, `TemplateValues`, `[OrdTemplate]`, `[TemplateName]`, `TemplateResult.Missing` |
| Out | `ToPlainText`, `ToMarkdown`, `SaveMarkdownAsync` |
| Macros | `HasMacros`, `RemoveMacros` (macros are kept or removed, never run) |
| Active content | `ActiveContent`, `ExternalReferences`, `RemoveActiveContent`; `DocxReader.GetActiveContentAsync`, `GetExternalReferencesAsync` |
| Limits and errors | `OrdLimits`, `ErrorKind()`, `OrdErrorKind` |

## License

MIT
