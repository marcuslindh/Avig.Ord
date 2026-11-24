using System.Text;
using System.Text.RegularExpressions;

namespace Avig.Ord.WordCheck;

/// <summary>A paragraph as a reader sees it: its list label (empty when it is not a list item) and its text.</summary>
/// <param name="Label">The list label, such as "1." or "•".</param>
/// <param name="Text">The text, with '\n' for a line, page or column break and '\t' for a tab.</param>
internal sealed record ParagraphText(string Label, string Text);

/// <summary>
/// The text of a document, as Word or Avig.Ord reads it, in the form both can be compared in:
/// the body's paragraphs with their list labels (empty paragraphs left out, since Word and the
/// package disagree on where a table puts them), and the text of the other stories.
/// </summary>
internal sealed record Report
{
    /// <summary>The body's paragraphs as they are, with tracked changes shown (Word's "All Markup", the package's Marked view).</summary>
    public List<ParagraphText> Body { get; init; } = [];

    /// <summary>The body's paragraphs with every change accepted.</summary>
    public List<ParagraphText> Accepted { get; init; } = [];

    /// <summary>How many pictures the body has in its text (inline pictures; an SVG with its PNG is one).</summary>
    public int Pictures { get; init; }

    /// <summary>The non-empty lines of the headers.</summary>
    public List<string> Headers { get; init; } = [];

    /// <summary>The non-empty lines of the footers.</summary>
    public List<string> Footers { get; init; } = [];

    /// <summary>The text of each footnote.</summary>
    public List<string> Footnotes { get; init; } = [];

    /// <summary>The text of each endnote.</summary>
    public List<string> Endnotes { get; init; } = [];

    /// <summary>The text of each comment.</summary>
    public List<string> Comments { get; init; } = [];

    /// <summary>The non-empty lines of the text boxes.</summary>
    public List<string> TextBoxes { get; init; } = [];
}

/// <summary>Turns Word's Range.Text into the package's text.</summary>
internal static partial class WordText
{
    /// <summary>
    /// Word writes a paragraph's text with control characters of its own: \r ends a paragraph,
    /// \a a table cell, \v is a line break, \f a page break, 14 a column break, 30 a
    /// non-breaking hyphen and 31 an optional one, 2 a note's reference mark, 5 a comment's,
    /// 1 a picture, 19–21 a field's code markers. The package has '\n' for every break and
    /// U+2011 for the non-breaking hyphen, and nothing for the rest.
    /// </summary>
    public static string Normalize(string text)
    {
        StringBuilder normalized = new(text.Length);

        foreach (char character in text)
        {
            switch (character)
            {
                case '\n':
                case '\v':
                case '\f':
                case (char)14:
                    normalized.Append('\n');

                    break;

                case (char)30:
                    normalized.Append('‑');

                    break;

                case '\t':
                    normalized.Append('\t');

                    break;

                case < ' ':
                    break;

                default:
                    normalized.Append(character);

                    break;
            }
        }

        return normalized.ToString();
    }

    /// <summary>
    /// Lines with every number as "#". Word works PAGE and NUMPAGES out again when it lays a
    /// document out; the package has the result that was saved. Headers and footers are
    /// compared this way.
    /// </summary>
    public static List<string> MaskNumbers(IEnumerable<string> lines)
    {
        return [.. lines.Select(line => Numbers().Replace(line, "#"))];
    }

    /// <summary>
    /// The package's text in the form Word gives it, for what cannot be mapped the other way:
    /// Word gives a symbol (w:sym, in the private-use area of a symbol font) as "(".
    /// </summary>
    public static string AsWordGivesIt(string text)
    {
        StringBuilder mapped = new(text.Length);

        foreach (char character in text)
        {
            if (character is >= '' and <= '')
            {
                mapped.Append('(');
            }
            else
            {
                mapped.Append(character);
            }
        }

        return mapped.ToString();
    }

    /// <summary>The non-empty lines of a story's text (a header, a text box): Word ends paragraphs with \r, the package with '\n'.</summary>
    public static List<string> Lines(string text)
    {
        return [.. Normalize(text.Replace('\r', '\n')).Split('\n').Select(line => line.Trim()).Where(line => line.Length > 0)];
    }

    [GeneratedRegex("[0-9]+")]
    private static partial Regex Numbers();
}
