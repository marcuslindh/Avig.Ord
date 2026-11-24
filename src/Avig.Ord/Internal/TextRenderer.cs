using System.Text;

namespace Avig.Ord.Internal;

/// <summary>How <see cref="TextRenderer"/> writes plain text.</summary>
internal sealed record TextOptions
{
    public static TextOptions Default { get; } = new();

    /// <summary>Put a list item's label (1., a), •) before its text.</summary>
    public bool ListLabels { get; init; } = true;

    /// <summary>Write the text of text boxes after the paragraph they are anchored in.</summary>
    public bool TextBoxes { get; init; } = true;
}

/// <summary>
/// Plain text: each paragraph on a line of its own, a list item with its label, a table row on
/// a line with its cells separated by tabs (a cell's own line breaks become spaces, so that a
/// row stays one line), and the text of text boxes after their paragraph.
/// </summary>
internal static class TextRenderer
{
    public static string Render(IEnumerable<Block> blocks, TextOptions options)
    {
        StringBuilder text = new();
        Write(text, blocks, options, 0);

        if (text.Length > 0 && text[^1] == '\n')
        {
            text.Length--;
        }

        return text.ToString();
    }

    public static void Write(StringBuilder text, IEnumerable<Block> blocks, TextOptions options, int depth)
    {
        foreach (Block block in blocks)
        {
            if (block is Paragraph paragraph)
            {
                WriteParagraph(text, paragraph, options, depth);
            }
            else if (block is Table table)
            {
                WriteTable(text, table, options, depth);
            }
        }
    }

    private static void WriteParagraph(StringBuilder text, Paragraph paragraph, TextOptions options, int depth)
    {
        IReadOnlyList<Inline> inlines = paragraph.Inlines;
        ListInfo? list = null;

        if (options.ListLabels == true)
        {
            list = paragraph.List;
        }

        if (list is not null && list.Label.Length > 0)
        {
            text.Append(list.Label).Append(' ');
        }

        foreach (Inline inline in inlines)
        {
            text.Append(inline.Text);
        }

        text.Append('\n');

        if (options.TextBoxes == true && depth < 8)
        {
            foreach (TextBox box in InlineReader.Flatten(inlines).OfType<TextBox>())
            {
                Write(text, box.Blocks, options, depth + 1);
            }
        }
    }

    private static void WriteTable(StringBuilder text, Table table, TextOptions options, int depth)
    {
        foreach (TableRow row in table.Rows)
        {
            bool first = true;

            foreach (TableCell cell in row.Cells)
            {
                if (first == false)
                {
                    text.Append('\t');
                }

                first = false;

                if (cell.VerticalMerge == VerticalMerge.Continue || depth >= 8)
                {
                    continue;
                }

                StringBuilder inner = new();
                Write(inner, cell.Blocks, options, depth + 1);
                text.Append(inner.ToString().TrimEnd('\n').Replace('\n', ' '));
            }

            text.Append('\n');
        }
    }
}
