using System.Globalization;
using System.Text;

namespace Avig.Ord.Internal;

/// <summary>
/// GitHub Flavored Markdown. A paragraph with an outline level is a heading (# to ######), a
/// list item a - or 1. item indented four spaces a level, a table a GFM table with its first row
/// as the header, bold, italic and strike-through ** * ~~, a hyperlink [text](url), and a picture
/// ![alt](link). Footnotes, comments, headers and footers are left out.
/// </summary>
internal sealed class MarkdownRenderer
{
    private readonly MarkdownOptions options;
    private readonly StringBuilder output = new();
    private bool lastWasList;

    private MarkdownRenderer(MarkdownOptions options)
    {
        this.options = options;
    }

    public static string Render(IEnumerable<Block> blocks, MarkdownOptions options)
    {
        MarkdownRenderer renderer = new(options);
        renderer.WriteBlocks(blocks, 0);

        return renderer.output.ToString().TrimEnd('\n') + "\n";
    }

    /// <summary>A renderer for a body read a block at a time: <see cref="Next"/> gives what each block adds, and <see cref="Finish"/> the end. Together they give what <see cref="Render"/> gives.</summary>
    public static MarkdownRenderer ForStream(MarkdownOptions options)
    {
        return new MarkdownRenderer(options);
    }

    /// <summary>The Markdown one more block adds. The blank lines after it are held back until the next block, since the last block has none.</summary>
    public string Next(Block block)
    {
        WriteBlocks([block], 0);
        int end = output.Length;

        while (end > 0 && output[end - 1] == '\n')
        {
            end--;
        }

        string written = output.ToString(0, end);
        output.Remove(0, end);

        return written;
    }

    /// <summary>The end of the Markdown: the one line end <see cref="Render"/> ends with.</summary>
    public static string Finish()
    {
        return "\n";
    }

    /// <summary>The file name a picture is linked as: its part's file name (image1.png), or a made-up one for a linked picture.</summary>
    public static string PictureFileName(Picture picture)
    {
        string part = picture.PartName;

        if (part.Length == 0)
        {
            return "picture";
        }

        return part[(part.LastIndexOf('/') + 1)..];
    }

    private static string Escape(string text, bool inTable)
    {
        StringBuilder escaped = new(text.Length);

        foreach (char character in text)
        {
            if (character is '\\' or '`' or '*' or '_' or '[' or ']' or '<' or '>' or '~' || (inTable == true && character == '|'))
            {
                escaped.Append('\\');
            }

            escaped.Append(character);
        }

        return escaped.ToString();
    }

    /// <summary>A line that would read as a heading, list item, quote or rule escaped at its start.</summary>
    private static string EscapeLineStart(string line)
    {
        string trimmed = line.TrimStart();

        if (trimmed.Length == 0)
        {
            return line;
        }

        if (trimmed[0] is '#' or '+' or '-' or '=')
        {
            return "\\" + trimmed;
        }

        int digits = 0;

        while (digits < trimmed.Length && char.IsAsciiDigit(trimmed[digits]) == true)
        {
            digits++;
        }

        if (digits > 0 && digits < trimmed.Length && trimmed[digits] is '.' or ')')
        {
            return trimmed[..digits] + "\\" + trimmed[digits..];
        }

        return trimmed;
    }

    private static string Wrap(string text, string marker)
    {
        // Emphasis may not start or end with a space: move the spaces outside the markers.
        string trimmed = text.Trim(' ');

        if (trimmed.Length == 0)
        {
            return text;
        }

        int leading = text.Length - text.TrimStart(' ').Length;
        int trailing = text.Length - text.TrimEnd(' ').Length;

        return new string(' ', leading) + marker + trimmed + marker + new string(' ', trailing);
    }

    private void WriteBlocks(IEnumerable<Block> blocks, int depth)
    {
        foreach (Block block in blocks)
        {
            if (block is Paragraph paragraph)
            {
                WriteParagraph(paragraph, depth);
            }
            else if (block is Table table)
            {
                WriteTable(table);
            }
        }
    }

    private void WriteParagraph(Paragraph paragraph, int depth)
    {
        IReadOnlyList<Inline> inlines = paragraph.Inlines;
        string text = Inlines(inlines, false).Trim();
        ListInfo? list = paragraph.List;
        int? heading = paragraph.OutlineLevel;

        if (text.Length > 0)
        {
            if (heading is not null)
            {
                Separate(false);
                output.Append('#', Math.Min(heading.Value + 1, 6)).Append(' ').Append(text.Replace("  \n", " ", StringComparison.Ordinal)).Append("\n\n");
            }
            else if (list is not null)
            {
                Separate(true);
                string marker = "-";

                if (list.IsBullet == false)
                {
                    marker = list.Number.ToString(CultureInfo.InvariantCulture) + ".";
                }

                string indent = new(' ', list.Level * 4);
                output.Append(indent).Append(marker).Append(' ').Append(text.Replace("\n", "\n" + indent + "    ", StringComparison.Ordinal)).Append('\n');
            }
            else
            {
                Separate(false);
                output.Append(EscapeLineStart(text)).Append("\n\n");
            }
        }

        if (options.TextBoxes == true && depth < 8)
        {
            foreach (TextBox box in InlineReader.Flatten(inlines).OfType<TextBox>())
            {
                WriteBlocks(box.Blocks, depth + 1);
            }
        }
    }

    /// <summary>A blank line before a block that is not the next item of the list before it.</summary>
    private void Separate(bool listItem)
    {
        if (lastWasList == true && listItem == false)
        {
            output.Append('\n');
        }

        lastWasList = listItem;
    }

    private void WriteTable(Table table)
    {
        IReadOnlyList<TableRow> rows = table.Rows;

        if (rows.Count == 0)
        {
            return;
        }

        Separate(false);

        // No more than Word allows: a w:tblGrid of a hundred thousand w:gridCol padded every row
        // to that many cells, a few megabytes of file to gigabytes of Markdown.
        int columns = Math.Min(Math.Max(table.ColumnCount, rows.Max(row => Table.SpannedColumns(row))), TableCell.MaxColumns);
        bool header = true;

        foreach (TableRow row in rows)
        {
            List<string> cells = [];

            foreach (TableCell cell in row.Cells)
            {
                string content = string.Empty;

                if (cell.VerticalMerge != VerticalMerge.Continue)
                {
                    content = CellContent(cell);
                }

                cells.Add(content);

                for (int span = 1; span < cell.ColumnSpan; span++)
                {
                    cells.Add(string.Empty);
                }
            }

            while (cells.Count < columns)
            {
                cells.Add(string.Empty);
            }

            output.Append("| ").Append(string.Join(" | ", cells)).Append(" |\n");

            if (header == true)
            {
                output.Append('|').Append(string.Concat(Enumerable.Repeat(" --- |", columns))).Append('\n');
                header = false;
            }
        }

        output.Append('\n');
    }

    /// <summary>A cell on one line: its paragraphs separated by &lt;br&gt;, a table inside it as its text.</summary>
    private string CellContent(TableCell cell)
    {
        List<string> lines = [];

        foreach (Block block in cell.Blocks)
        {
            string line;

            if (block is Paragraph paragraph)
            {
                line = Inlines(paragraph.Inlines, true).Trim();
                ListInfo? list = paragraph.List;

                if (list is not null && line.Length > 0)
                {
                    line = Escape(list.Label, true) + " " + line;
                }
            }
            else
            {
                line = Escape(block.Text.Replace('\n', ' ').Replace('\t', ' '), true);
            }

            if (line.Length > 0)
            {
                lines.Add(line.Replace("  \n", "<br>", StringComparison.Ordinal).Replace("\n", "<br>", StringComparison.Ordinal));
            }
        }

        return string.Join("<br>", lines);
    }

    private string Inlines(IReadOnlyList<Inline> inlines, bool inTable)
    {
        StringBuilder text = new();
        List<(string Text, RunFormat Format)> pending = [];

        foreach (Inline inline in inlines)
        {
            if (inline is Run run)
            {
                pending.Add((run.Text, run.Format));

                continue;
            }

            Flush(text, pending, inTable);

            switch (inline)
            {
                case Hyperlink link:
                    WriteLink(text, Plain(link.Inlines), link.Url, link.Anchor, inTable);

                    break;

                case Field field when field.Code == "HYPERLINK":
                    WriteLink(text, Plain(field.Result), FieldInstruction.FirstArgument(field.Instruction), null, inTable);

                    break;

                case Field field:
                    text.Append(Inlines(field.Result, inTable));

                    break;

                case ContentControl control:
                    text.Append(Inlines(control.Inlines, inTable));

                    break;

                case Picture picture:
                    text.Append("![").Append(Escape(picture.Description, inTable)).Append("](").Append(PictureLink(picture)).Append(')');

                    break;

                case Break:
                    text.Append(' ');

                    break;

                default:
                    break;
            }
        }

        Flush(text, pending, inTable);

        return text.ToString();
    }

    private static void WriteLink(StringBuilder text, string label, string? url, string? anchor, bool inTable)
    {
        string target = url ?? string.Empty;

        if (target.Length == 0 && anchor is not null)
        {
            target = "#" + anchor;
        }

        if (target.Length == 0)
        {
            text.Append(Escape(label, inTable));

            return;
        }

        text.Append('[').Append(Escape(label, inTable)).Append("](").Append(target.Replace(" ", "%20", StringComparison.Ordinal).Replace(")", "%29", StringComparison.Ordinal)).Append(')');
    }

    private string PictureLink(Picture picture)
    {
        if (options.PictureLink is not null)
        {
            return options.PictureLink(picture);
        }

        string folder = options.PictureFolder.TrimEnd('/');
        string name = PictureFileName(picture);

        if (folder.Length == 0)
        {
            return name;
        }

        return folder + "/" + name;
    }

    private static string Plain(IReadOnlyList<Inline> inlines)
    {
        return string.Concat(inlines.Select(inline => inline.Text)).Replace('\n', ' ');
    }

    /// <summary>Writes the runs gathered so far, with runs of the same emphasis joined so that markers are not repeated between them.</summary>
    private static void Flush(StringBuilder text, List<(string Text, RunFormat Format)> pending, bool inTable)
    {
        int index = 0;

        while (index < pending.Count)
        {
            (bool bold, bool italic, bool strike) = Emphasis(pending[index].Format);
            StringBuilder joined = new();

            while (index < pending.Count && Emphasis(pending[index].Format) == (bold, italic, strike))
            {
                joined.Append(pending[index].Text);
                index++;
            }

            string segment = Escape(joined.ToString(), inTable).Replace("\t", " ", StringComparison.Ordinal).Replace("\n", "  \n", StringComparison.Ordinal);

            if (strike == true)
            {
                segment = Wrap(segment, "~~");
            }

            if (italic == true)
            {
                segment = Wrap(segment, "*");
            }

            if (bold == true)
            {
                segment = Wrap(segment, "**");
            }

            text.Append(segment);
        }

        pending.Clear();
    }

    private static (bool Bold, bool Italic, bool Strike) Emphasis(RunFormat format)
    {
        return (format.Bold == true, format.Italic == true, format.Strike == true);
    }
}
