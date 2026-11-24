using System.Text.RegularExpressions;

namespace Avig.Ord;

/// <summary>
/// What a match is replaced with: text (in the formatting the match had, or another), a
/// picture, or whole paragraphs and tables, which split the paragraph the match is in.
/// </summary>
/// <example>
/// <code>
/// document.Replace("{{Logo}}", Replacement.Picture(logo, "Company logo"));
/// document.Replace("{{Prices}}", Replacement.Table(t =&gt; t.HeaderRow("Item", "Price").Row("Tea", "25")));
/// </code>
/// </example>
public sealed class Replacement
{
    private Replacement(ReplacementKind kind)
    {
        Kind = kind;
    }

    internal ReplacementKind Kind { get; }

    internal string TextValue { get; private init; } = string.Empty;

    internal Action<RunStyle>? Style { get; private init; }

    internal byte[]? PictureData { get; private init; }

    internal Length? PictureWidth { get; private init; }

    internal Length? PictureHeight { get; private init; }

    internal string Description { get; private init; } = string.Empty;

    internal IReadOnlyList<Action<ParagraphBuilder>> ParagraphBuilders { get; private init; } = [];

    internal Action<TableBuilder>? TableBuilder { get; private init; }

    /// <summary>Text, in the formatting of the first run of the match. '\t' and '\n' become a tab and a line break.</summary>
    public static Replacement Text(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        return new Replacement(ReplacementKind.Text) { TextValue = text };
    }

    /// <summary>Text in a formatting of its own, set over the formatting the match had.</summary>
    public static Replacement Text(string text, Action<RunStyle> style)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(style);

        return new Replacement(ReplacementKind.Text) { TextValue = text, Style = style };
    }

    /// <summary>Nothing: the match is removed.</summary>
    public static Replacement Nothing { get; } = new(ReplacementKind.Text);

    /// <summary>A picture in the line, at its own size (no wider than the text).</summary>
    public static Replacement Picture(byte[] data, string description = "")
    {
        ArgumentNullException.ThrowIfNull(data);

        return new Replacement(ReplacementKind.Picture) { PictureData = data, Description = description };
    }

    /// <summary>A picture in the line at a given size.</summary>
    public static Replacement Picture(byte[] data, Length width, Length height, string description = "")
    {
        ArgumentNullException.ThrowIfNull(data);

        return new Replacement(ReplacementKind.Picture) { PictureData = data, PictureWidth = width, PictureHeight = height, Description = description };
    }

    /// <summary>
    /// Paragraphs: the paragraph the match is in is split there, and the new paragraphs go in
    /// between. A match that is all of its paragraph replaces the paragraph.
    /// </summary>
    public static Replacement Paragraphs(params Action<ParagraphBuilder>[] paragraphs)
    {
        ArgumentNullException.ThrowIfNull(paragraphs);

        return new Replacement(ReplacementKind.Blocks) { ParagraphBuilders = paragraphs };
    }

    /// <summary>Paragraphs of plain text, one for each string.</summary>
    public static Replacement Paragraphs(IEnumerable<string> paragraphs)
    {
        ArgumentNullException.ThrowIfNull(paragraphs);

        return Paragraphs([.. paragraphs.Select(text => (Action<ParagraphBuilder>)(paragraph => paragraph.Text(text)))]);
    }

    /// <summary>A table: the paragraph the match is in is split there, and the table goes in between.</summary>
    public static Replacement Table(Action<TableBuilder> table)
    {
        ArgumentNullException.ThrowIfNull(table);

        return new Replacement(ReplacementKind.Blocks) { TableBuilder = table };
    }

    /// <summary>Text, as <see cref="Text(string)"/>.</summary>
    public static implicit operator Replacement(string text)
    {
        return Text(text);
    }
}

internal enum ReplacementKind
{
    Text,
    Picture,
    Blocks,
}

/// <summary>Which parts of a document search and replace go through.</summary>
[Flags]
public enum SearchScope
{
    /// <summary>The body, with its tables, text boxes and content controls.</summary>
    Body = 1,

    /// <summary>Every header and footer.</summary>
    HeadersAndFooters = 2,

    /// <summary>Footnotes and endnotes.</summary>
    Notes = 4,

    /// <summary>Comments.</summary>
    Comments = 8,

    /// <summary>All of them: the default.</summary>
    All = Body | HeadersAndFooters | Notes | Comments,
}

/// <summary>How text is searched for.</summary>
public sealed record SearchOptions
{
    /// <summary>The options used when none are given: case matters, anywhere in a word, everywhere.</summary>
    public static SearchOptions Default { get; } = new();

    /// <summary>Whether upper and lower case must match. True by default. (For a regex, its own options decide.)</summary>
    public bool MatchCase { get; init; } = true;

    /// <summary>Only matches that are whole words: no letter or digit right before or after.</summary>
    public bool WholeWord { get; init; }

    /// <summary>Which parts are searched. <see cref="SearchScope.All"/> by default.</summary>
    public SearchScope Scope { get; init; } = SearchScope.All;
}

/// <summary>A match found by <see cref="OrdDocument.Find(string, SearchOptions?)"/>.</summary>
/// <param name="Paragraph">The paragraph it is in. A match never goes past its paragraph.</param>
/// <param name="Index">Where it starts in the paragraph's text.</param>
/// <param name="Value">The text matched.</param>
/// <param name="RegexMatch">The regex match, with its groups, when the search was a regex.</param>
public sealed record TextMatch(Paragraph Paragraph, int Index, string Value, Match? RegexMatch);
