using System.Text;
using System.Xml;

namespace Avig.Ord.Fuzz;

/// <summary>
/// What every file target does with a file: read it with both readers in every way the package
/// offers, change it, save it, and open what was saved. A damaged file must be refused with a
/// Result failure that has an <see cref="OrdErrorKind"/>, or, while <see cref="DocxReader"/> reads
/// the body, with the <see cref="InvalidDataException"/> or <see cref="XmlException"/> it documents.
/// Once <see cref="OrdDocument.OpenAsync(Stream, ReadOptions?, CancellationToken)"/> has succeeded,
/// nothing may throw: the file was accepted.
/// </summary>
internal static class Harness
{
    private static readonly RevisionView[] Views = [RevisionView.Accepted, RevisionView.Original, RevisionView.Marked];

    /// <summary>How many refusals the targets have seen, for the smoke run to count inputs as refused rather than passed.</summary>
    public static int Refusals { get; private set; }

    /// <summary>A refused file must have a kind an application can show its users.</summary>
    public static void CheckKind(OrdErrorKind kind, string? error)
    {
        Refusals++;

        if (kind == OrdErrorKind.Unknown)
        {
            throw new InvariantViolationException($"A refusal has no error kind: {error}");
        }
    }

    /// <summary>Reads, changes, saves and reopens a file; the view is picked by <paramref name="selector"/>.</summary>
    public static async Task RunFileAsync(byte[] file, int selector)
    {
        RevisionView view = Views[Math.Abs(selector % Views.Length)];
        await ReadStreamingAsync(file, view);
        await EditAndSaveAsync(file, view, (selector & 4) != 0);
    }

    /// <summary>Reads the body with <see cref="DocxReader"/> three times: as blocks, as events, and as Markdown and text.</summary>
    public static async Task ReadStreamingAsync(byte[] file, RevisionView view)
    {
        ReadOptions options = new() { Revisions = view };

        await StreamAsync(file, options, async reader =>
        {
            await foreach (Block block in reader.ReadBlocksAsync())
            {
                await TouchAsync(block);
            }

            foreach (Style style in reader.Styles.Styles)
            {
                _ = reader.Styles.TryGet(style.Id, out _);
            }
        });

        await StreamAsync(file, options, async reader =>
        {
            await foreach (DocumentEvent item in reader.ReadEventsAsync())
            {
                _ = item.Text.Length;
            }
        });

        await StreamAsync(file, options, async reader =>
        {
            await using StringWriter markdown = new();
            await reader.WriteMarkdownAsync(markdown);
        });

        await StreamAsync(file, options, async reader => _ = await reader.ReadPlainTextAsync());
    }

    /// <summary>Opens a streaming reader and runs <paramref name="read"/>; a refusal, or a damaged part found while reading, must say what kind of failure it is.</summary>
    public static async Task StreamAsync(byte[] file, ReadOptions options, Func<DocxReader, Task> read)
    {
        Result<DocxReader> opened = await DocxReader.OpenAsync(new MemoryStream(file, false), options);

        if (opened.IsSuccess(out DocxReader? reader) == false)
        {
            CheckKind(opened.ErrorKind(), opened.Error);

            return;
        }

        await using (reader)
        {
            try
            {
                await read(reader);
            }
            catch (Exception exception) when (exception is InvalidDataException or XmlException)
            {
                // Documented on ReadBlocksAsync: a part that turns out to be damaged while it is read.
                CheckKind(exception.ErrorKind(), exception.Message);
            }
        }
    }

    /// <summary>
    /// Opens the file as a document, reads everything in it, changes it and saves it. The saved
    /// file must open again, hold the change, and read through without being called damaged:
    /// the package wrote it.
    /// </summary>
    public static async Task EditAndSaveAsync(byte[] file, RevisionView view, bool flat)
    {
        Result<OrdDocument> opened = await OrdDocument.OpenAsync(new MemoryStream(file, false), new ReadOptions { Revisions = view });

        if (opened.IsSuccess(out OrdDocument? document) == false)
        {
            CheckKind(opened.ErrorKind(), opened.Error);

            return;
        }

        await ExploreAsync(document);
        Edit(document);
        await ExploreAsync(document);
        byte[] saved = await SaveAsync(document, new SaveOptions { FlatOpc = flat });

        try
        {
            Result<OrdDocument> reopened = await OrdDocument.OpenAsync(new MemoryStream(saved, false), new ReadOptions { Revisions = view });

            if (reopened.IsSuccess(out OrdDocument? again) == false)
            {
                throw new InvariantViolationException($"A saved document cannot be opened again: {reopened.Error}");
            }

            await ExploreAsync(again);
            string text = again.ToPlainText();

            if (text.Contains("changed", StringComparison.Ordinal) == false || text.Contains("added cell", StringComparison.Ordinal) == false)
            {
                throw new InvariantViolationException($"The change is not in the saved document: {Shorten(text)}");
            }

            if (again.Properties.Title != "Fuzz")
            {
                throw new InvariantViolationException($"The title reads back as '{again.Properties.Title}'.");
            }

            await ReadStreamingAsync(saved, view);
        }
        catch (Exception exception) when (exception is InvalidDataException or XmlException)
        {
            throw new InvariantViolationException("A document the package saved is called damaged by the package.", exception);
        }
    }

    /// <summary>The changes every file target makes: text, a list item, a table, a picture, a note and a property.</summary>
    public static void Edit(OrdDocument document)
    {
        IReadOnlyList<Paragraph> paragraphs = document.Paragraphs;

        if (paragraphs.Count > 0)
        {
            paragraphs[0].AppendText(" appended", style => style.Bold());
            paragraphs[^1].InsertParagraphAfter("after the last");
        }

        if (document.Tables.Count > 0)
        {
            document.Tables[0].AddRow("x", "y");
        }

        document.AddParagraph(paragraph => paragraph.Bullet().Text("changed").Footnote("note").Link(" link", "https://example.com/"));
        document.AddParagraph(paragraph => paragraph.Numbered(1).Picture(Template.Png, "picture"));
        document.AddTable(table => table.HeaderRow("added cell", "b").Row("c", "d"));
        document.AddHeading("Heading", 2);
        document.Properties.Title = "Fuzz";
    }

    /// <summary>How many Flat OPC saves hit the known depth limit of Avig.OpenXml's writer (see <see cref="SaveAsync"/>).</summary>
    public static int KnownFlatDepthIssues { get; private set; }

    public static async Task<byte[]> SaveAsync(OrdDocument document, SaveOptions options)
    {
        try
        {
            using MemoryStream output = new();
            await document.SaveAsync(output, options);

            return output.ToArray();
        }
        catch (XmlException exception) when (options.FlatOpc == true && exception.Message.Contains("nested more than 100 levels", StringComparison.Ordinal) == true)
        {
            // Known, in Avig.OpenXml 0.1.2: OpcPackageWriter re-reads every XML part for Flat OPC
            // with its own depth limit of 100, while OrdLimits.MaxXmlDepth lets a document of 256
            // levels open. It has to be fixed there; FuzzFindingsTests has the test, skipped until
            // then. Saved as a zip instead, so that the rest of the harness still runs.
            KnownFlatDepthIssues++;

            return await SaveAsync(document, options with { FlatOpc = false });
        }
    }

    /// <summary>Everything the document model can say about a document.</summary>
    public static async Task ExploreAsync(OrdDocument document)
    {
        foreach (Block block in document.Blocks)
        {
            await TouchAsync(block);
        }

        _ = document.Paragraphs.Count;
        _ = document.Tables.Count;
        _ = document.ToPlainText();
        _ = document.ToMarkdown();
        _ = document.ToMarkdown(new MarkdownOptions { TextBoxes = false, PictureFolder = string.Empty, PictureLink = picture => picture.PartName });
        _ = document.Kind;
        _ = document.HasMacros;

        foreach (HeaderFooter part in document.Headers.Concat(document.Footers))
        {
            _ = part.Text;
            _ = part.Kind;

            foreach (Block block in part.Blocks)
            {
                await TouchAsync(block);
            }
        }

        foreach (Note note in document.Footnotes.Concat(document.Endnotes))
        {
            _ = note.Text + note.Id;
        }

        foreach (Comment comment in document.Comments)
        {
            _ = comment.Text + comment.Author + comment.Initials + comment.Id;
            _ = comment.Date;
        }

        foreach (Section section in document.Sections)
        {
            _ = section.PageWidth;
            _ = section.PageHeight;
            _ = section.Orientation;
            _ = section.MarginTop;
            _ = section.MarginBottom;
            _ = section.MarginLeft;
            _ = section.MarginRight;
        }

        DocumentProperties properties = document.Properties;
        _ = properties.Title + properties.Subject + properties.Author + properties.Keywords + properties.Description + properties.LastModifiedBy;
        _ = properties.Created;
        _ = properties.Modified;

        StyleSheet styles = document.Styles;
        _ = styles.DefaultRun;
        _ = styles.DefaultParagraph;

        foreach (Style style in styles.Styles)
        {
            _ = styles.TryFindByName(style.Name, out _);
        }
    }

    /// <summary>Reads every property of a block, and of everything in it.</summary>
    public static async Task TouchAsync(Block block)
    {
        _ = block.Text;

        if (block is Paragraph paragraph)
        {
            _ = paragraph.StyleId;
            _ = paragraph.StyleName;
            _ = paragraph.DirectFormat;
            _ = paragraph.Format;
            _ = paragraph.OutlineLevel;
            _ = paragraph.IsHeading;
            _ = paragraph.EndsSection;
            _ = paragraph.List?.Label;

            foreach (Run run in paragraph.Runs)
            {
                _ = run.Format;
                _ = run.DirectFormat;
                _ = run.StyleId;
                _ = run.Revision;
            }

            await TouchAsync(paragraph.Inlines);
        }
        else if (block is Table table)
        {
            _ = table.ColumnCount;
            _ = table.StyleId;

            foreach (TableRow row in table.Rows)
            {
                _ = row.IsHeader;
                _ = row.Text;

                foreach (TableCell cell in row.Cells)
                {
                    _ = cell.ColumnSpan;
                    _ = cell.VerticalMerge;
                    _ = cell.Paragraphs.Count;

                    foreach (Block inner in cell.Blocks)
                    {
                        await TouchAsync(inner);
                    }
                }
            }
        }
    }

    private static async Task TouchAsync(IReadOnlyList<Inline> inlines)
    {
        foreach (Inline inline in inlines)
        {
            _ = inline.Text;

            switch (inline)
            {
                case Hyperlink link:
                    _ = link.Url;
                    _ = link.Anchor;
                    await TouchAsync(link.Inlines);

                    break;

                case Field field:
                    _ = field.Code + field.Instruction + field.MergeFieldName;
                    await TouchAsync(field.Result);

                    break;

                case ContentControl control:
                    _ = control.Tag + control.Title;
                    await TouchAsync(control.Inlines);

                    break;

                case Picture picture:
                    _ = picture.PartName + picture.SvgPartName + picture.Description + picture.Title;
                    _ = picture.Width + picture.Height;
                    _ = await picture.ReadAsync();

                    break;

                case TextBox box:
                    foreach (Block block in box.Blocks)
                    {
                        await TouchAsync(block);
                    }

                    break;

                case NoteReference note:
                    _ = note.Kind;
                    _ = note.Id;

                    break;

                case Break pageBreak:
                    _ = pageBreak.Kind;

                    break;

                default:
                    break;
            }
        }
    }

    public static string Shorten(string text)
    {
        string line = text.ReplaceLineEndings("\\n");

        if (line.Length <= 300)
        {
            return line;
        }

        return line[..300] + "...";
    }

    /// <summary>The text of a list of strings, for a message: each quoted, one per line.</summary>
    public static string Lines(IEnumerable<string> lines)
    {
        StringBuilder text = new();

        foreach (string line in lines.Take(20))
        {
            text.Append('\n').Append("    \"").Append(Shorten(line)).Append('"');
        }

        return text.ToString();
    }
}
