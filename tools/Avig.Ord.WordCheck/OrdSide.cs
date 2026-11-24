namespace Avig.Ord.WordCheck;

/// <summary>What Avig.Ord reads in a document, in the same form as <see cref="WordSession.Read"/>, and how the two are compared.</summary>
internal static class OrdSide
{
    public static async Task<Report> ReadAsync(string path)
    {
        OrdDocument marked = await OpenAsync(path, RevisionView.Marked);
        OrdDocument accepted = await OpenAsync(path, RevisionView.Accepted);

        return new Report
        {
            Body = Body(marked),
            Accepted = Body(accepted),
            Pictures = marked.Paragraphs.Sum(paragraph => Flatten(paragraph.Inlines).OfType<Picture>().Count()),
            Headers = [.. marked.Headers.SelectMany(header => StoryLines(header.Blocks)).Select(line => WordText.MaskNumbers([line])[0]).Distinct().Order(StringComparer.Ordinal)],
            Footers = [.. marked.Footers.SelectMany(footer => StoryLines(footer.Blocks)).Select(line => WordText.MaskNumbers([line])[0]).Distinct().Order(StringComparer.Ordinal)],
            Footnotes = [.. marked.Footnotes.Select(note => string.Join('\n', WordText.Lines(note.Text)))],
            Endnotes = [.. marked.Endnotes.Select(note => string.Join('\n', WordText.Lines(note.Text)))],
            Comments = [.. marked.Comments.Select(comment => string.Join('\n', WordText.Lines(comment.Text)))],
            TextBoxes = [.. TextBoxes(marked.Blocks).SelectMany(WordText.Lines)],
        };
    }

    public static async Task<OrdDocument> OpenAsync(string path, RevisionView view)
    {
        Result<OrdDocument> opened = await OrdDocument.OpenAsync(path, new ReadOptions { Revisions = view });

        if (opened.IsSuccess(out OrdDocument? document) == false)
        {
            throw new InvalidOperationException($"Avig.Ord cannot open {path}: {opened.Error}");
        }

        return document;
    }

    /// <summary>The differences between what Word and Avig.Ord read, one line each; empty when they agree.</summary>
    public static List<string> Compare(Report word, Report ord)
    {
        List<string> differences = [];
        CompareParagraphs("body", word.Body, ord.Body, differences);

        if (word.Pictures != ord.Pictures)
        {
            differences.Add($"pictures: Word {word.Pictures}, Avig.Ord {ord.Pictures}");
        }

        CompareParagraphs("body with changes accepted", word.Accepted, ord.Accepted, differences);
        CompareLines("headers", word.Headers, ord.Headers, differences);
        CompareLines("footers", word.Footers, ord.Footers, differences);
        CompareLines("footnotes", word.Footnotes, ord.Footnotes, differences);
        CompareLines("endnotes", word.Endnotes, ord.Endnotes, differences);
        CompareLines("comments", word.Comments, ord.Comments, differences);
        CompareLines("text boxes", word.TextBoxes, ord.TextBoxes, differences);

        return differences;
    }

    /// <summary>The lines of a header or footer, with each field's result as "#" (see WordSession.HeadersAndFooters).</summary>
    private static List<string> StoryLines(IEnumerable<Block> blocks)
    {
        List<string> lines = [];

        foreach (Block block in blocks)
        {
            if (block is Paragraph paragraph)
            {
                lines.AddRange(WordText.Lines(string.Concat(paragraph.Inlines.Select(Masked))));
            }
            else if (block is Table table)
            {
                lines.AddRange(table.Rows.SelectMany(row => row.Cells).SelectMany(cell => StoryLines(cell.Blocks)));
            }
        }

        return lines;
    }

    private static string Masked(Inline inline)
    {
        return inline switch
        {
            Field => "#",
            Hyperlink link => string.Concat(link.Inlines.Select(Masked)),
            ContentControl control => string.Concat(control.Inlines.Select(Masked)),
            _ => inline.Text,
        };
    }

    private static List<ParagraphText> Body(OrdDocument document)
    {
        List<ParagraphText> paragraphs = [];

        foreach (Paragraph paragraph in document.Paragraphs)
        {
            if (paragraph.Text.Length == 0)
            {
                continue;
            }

            string label = paragraph.List?.Label ?? string.Empty;

            if (paragraph.List?.IsBullet == true)
            {
                label = "bullet";
            }

            paragraphs.Add(new ParagraphText(label, WordText.AsWordGivesIt(paragraph.Text)));
        }

        return paragraphs;
    }

    /// <summary>The text of every text box in the blocks, also in tables; a text box's own text boxes are its business.</summary>
    private static List<string> TextBoxes(IEnumerable<Block> blocks)
    {
        List<string> texts = [];

        foreach (Block block in blocks)
        {
            if (block is Table table)
            {
                foreach (TableCell cell in table.Rows.SelectMany(row => row.Cells))
                {
                    texts.AddRange(TextBoxes(cell.Blocks));
                }

                continue;
            }

            if (block is Paragraph paragraph)
            {
                foreach (TextBox box in Flatten(paragraph.Inlines).OfType<TextBox>())
                {
                    texts.Add(string.Join('\n', box.Blocks.Select(inner => inner.Text)));
                }
            }
        }

        return texts;
    }

    private static IEnumerable<Inline> Flatten(IEnumerable<Inline> inlines)
    {
        foreach (Inline inline in inlines)
        {
            yield return inline;

            IReadOnlyList<Inline> children = inline switch
            {
                Hyperlink link => link.Inlines,
                Field field => field.Result,
                ContentControl control => control.Inlines,
                _ => [],
            };

            foreach (Inline child in Flatten(children))
            {
                yield return child;
            }
        }
    }

    private static void CompareParagraphs(string what, List<ParagraphText> word, List<ParagraphText> ord, List<string> differences)
    {
        int count = Math.Max(word.Count, ord.Count);

        for (int index = 0; index < count; index++)
        {
            ParagraphText? fromWord = word.ElementAtOrDefault(index);
            ParagraphText? fromOrd = ord.ElementAtOrDefault(index);

            if (fromWord != fromOrd)
            {
                differences.Add($"{what}, paragraph {index}: Word {Show(fromWord)}, Avig.Ord {Show(fromOrd)}");

                // After the first difference the rest are usually shifted; one says enough.
                return;
            }
        }
    }

    private static void CompareLines(string what, List<string> word, List<string> ord, List<string> differences)
    {
        if (word.SequenceEqual(ord) == false)
        {
            differences.Add($"{what}: Word [{string.Join(" | ", word.Select(Escape))}], Avig.Ord [{string.Join(" | ", ord.Select(Escape))}]");
        }
    }

    private static string Show(ParagraphText? paragraph)
    {
        if (paragraph is null)
        {
            return "(none)";
        }

        return $"[{Escape(paragraph.Label)}] \"{Escape(paragraph.Text)}\"";
    }

    public static string Escape(string text)
    {
        return text.Replace("\n", "\\n", StringComparison.Ordinal).Replace("\t", "\\t", StringComparison.Ordinal);
    }
}
