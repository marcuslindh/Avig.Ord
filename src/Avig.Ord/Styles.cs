using System.Diagnostics.CodeAnalysis;
using System.Xml.Linq;
using Avig.Ord.Internal;

namespace Avig.Ord;

/// <summary>What a style formats.</summary>
public enum StyleKind
{
    /// <summary>A paragraph style, such as Normal or Heading 1. It can set run formatting too.</summary>
    Paragraph,

    /// <summary>A character style, such as Strong, applied to runs.</summary>
    Character,

    /// <summary>A table style, such as Table Grid.</summary>
    Table,

    /// <summary>A numbering style, for lists.</summary>
    Numbering,
}

/// <summary>A style in the document's styles part.</summary>
/// <param name="Id">The id paragraphs and runs refer to it by, such as Heading1. In a document in another language the id is often the same and the name translated.</param>
/// <param name="Name">The name Word shows, such as "heading 1" (built-in styles have lower-case names in the file) or "Rubrik 1".</param>
/// <param name="Kind">What it formats.</param>
/// <param name="BasedOn">The id of the style it inherits from, if any.</param>
/// <param name="IsDefault">True for the default style of its kind: the one a paragraph without a style has.</param>
public sealed record Style(string Id, string Name, StyleKind Kind, string? BasedOn, bool IsDefault);

/// <summary>
/// The styles of a document, and how formatting is inherited: from the document defaults,
/// through a paragraph style and the styles it is based on, through a character style, to what
/// is set on the paragraph or run itself.
/// </summary>
public sealed class StyleSheet
{
    /// <summary>How long a basedOn chain may be. Word stops at a few dozen; a loop in a file stops here too.</summary>
    private const int MaxChain = 64;

    private readonly WordNames names;
    private readonly Dictionary<string, Entry> byId = new(StringComparer.Ordinal);
    private readonly Dictionary<string, (RunFormat Run, ParagraphFormat Paragraph)> paragraphCache = new(StringComparer.Ordinal);
    private readonly Dictionary<string, RunFormat> characterCache = new(StringComparer.Ordinal);
    private readonly List<Style> styles = [];

    private StyleSheet(WordNames names, XElement? root)
    {
        this.names = names;
        XElement? defaults = root?.Element(names.W + "docDefaults");
        DefaultRun = FormatReader.Run(names, defaults?.Element(names.W + "rPrDefault")?.Element(names.RunProperties));
        DefaultParagraph = FormatReader.Paragraph(names, defaults?.Element(names.W + "pPrDefault")?.Element(names.ParagraphProperties));

        foreach (XElement element in root?.Elements(names.Style) ?? [])
        {
            string? id = element.Attribute(names.StyleId)?.Value;

            if (id is null || byId.ContainsKey(id) == true)
            {
                continue;
            }

            StyleKind kind = element.Attribute(names.Type)?.Value switch
            {
                "character" => StyleKind.Character,
                "table" => StyleKind.Table,
                "numbering" => StyleKind.Numbering,
                _ => StyleKind.Paragraph,
            };

            bool isDefault = IsOn(element.Attribute(names.W + "default")?.Value);
            Style style = new(id, names.ValueOf(element, "name") ?? id, kind, names.ValueOf(element, "basedOn"), isDefault);
            byId[id] = new Entry(style, element);
            styles.Add(style);
        }

        DefaultParagraphStyle = styles.FirstOrDefault(style => style.Kind == StyleKind.Paragraph && style.IsDefault == true)?.Id;
        DefaultCharacterStyle = styles.FirstOrDefault(style => style.Kind == StyleKind.Character && style.IsDefault == true)?.Id;
    }

    /// <summary>Every style, in the order of the file.</summary>
    public IReadOnlyList<Style> Styles => styles;

    /// <summary>The run formatting every run starts from (w:docDefaults).</summary>
    public RunFormat DefaultRun { get; }

    /// <summary>The paragraph formatting every paragraph starts from (w:docDefaults).</summary>
    public ParagraphFormat DefaultParagraph { get; }

    /// <summary>The id of the paragraph style a paragraph without one has, usually Normal.</summary>
    public string? DefaultParagraphStyle { get; }

    internal string? DefaultCharacterStyle { get; }

    /// <summary>A style by its id.</summary>
    public bool TryGet(string id, [NotNullWhen(true)] out Style? style)
    {
        style = null;

        if (byId.TryGetValue(id, out Entry? entry) == true)
        {
            style = entry.Style;
        }

        return style is not null;
    }

    /// <summary>A style by the name Word shows, case ignored.</summary>
    public bool TryFindByName(string name, [NotNullWhen(true)] out Style? style)
    {
        style = styles.FirstOrDefault(candidate => string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase) == true);

        return style is not null;
    }

    internal static StyleSheet Read(XElement? root)
    {
        return new StyleSheet(WordNames.For(root), root);
    }

    /// <summary>
    /// The paragraph formatting a paragraph with this style (or the default style, for null)
    /// gets from the styles, and the run formatting the style gives its runs.
    /// </summary>
    internal (RunFormat Run, ParagraphFormat Paragraph) ParagraphStyle(string? styleId)
    {
        string id = styleId ?? DefaultParagraphStyle ?? string.Empty;

        if (paragraphCache.TryGetValue(id, out (RunFormat Run, ParagraphFormat Paragraph) cached) == true)
        {
            return cached;
        }

        RunFormat run = DefaultRun;
        ParagraphFormat paragraph = DefaultParagraph;

        foreach (XElement element in Chain(id))
        {
            run = run.With(FormatReader.Run(names, element.Element(names.RunProperties)));
            paragraph = paragraph.With(FormatReader.Paragraph(names, element.Element(names.ParagraphProperties)));
        }

        // The Heading 1–9 a file made by a program other than Word may have without an outline level.
        if (paragraph.OutlineLevel is null && TryGet(id, out Style? style) == true && TryHeadingNumber(style.Name, out int level) == true)
        {
            paragraph = paragraph with { OutlineLevel = level - 1 };
        }

        paragraphCache[id] = (run, paragraph);

        return (run, paragraph);
    }

    /// <summary>The run formatting a character style gives, through the styles it is based on.</summary>
    internal RunFormat CharacterStyle(string? styleId)
    {
        if (styleId is null)
        {
            return RunFormat.None;
        }

        if (characterCache.TryGetValue(styleId, out RunFormat? cached) == true)
        {
            return cached;
        }

        RunFormat run = RunFormat.None;

        foreach (XElement element in Chain(styleId))
        {
            run = run.With(FormatReader.Run(names, element.Element(names.RunProperties)));
        }

        characterCache[styleId] = run;

        return run;
    }

    /// <summary>The numbering a paragraph style gives (w:numPr in the style), through the styles it is based on.</summary>
    internal (string? NumberingId, int? Level) StyleNumbering(string? styleId)
    {
        string? numberingId = null;
        int? level = null;

        foreach (XElement element in Chain(styleId ?? DefaultParagraphStyle ?? string.Empty))
        {
            XElement? numbering = element.Element(names.ParagraphProperties)?.Element(names.W + "numPr");

            if (numbering is null)
            {
                continue;
            }

            numberingId = names.ValueOf(numbering, "numId") ?? numberingId;
            level = FormatReader.Integer(names.ValueOf(numbering, "ilvl")) ?? level;
        }

        return (numberingId, level);
    }

    private static bool IsOn(string? value)
    {
        return value is "1" or "true" or "on";
    }

    private static bool TryHeadingNumber(string name, out int level)
    {
        level = 0;

        return name.StartsWith("heading ", StringComparison.OrdinalIgnoreCase) == true
            && int.TryParse(name.AsSpan(8), out level) == true
            && level is >= 1 and <= 9;
    }

    /// <summary>The style and the styles it is based on, the base first. A loop or a missing style ends the chain.</summary>
    private List<XElement> Chain(string id)
    {
        List<XElement> chain = [];
        HashSet<string> seen = new(StringComparer.Ordinal);
        string? current = id;

        while (current is not null && chain.Count < MaxChain && seen.Add(current) == true && byId.TryGetValue(current, out Entry? entry) == true)
        {
            chain.Add(entry.Element);
            current = entry.Style.BasedOn;
        }

        chain.Reverse();

        return chain;
    }

    private sealed record Entry(Style Style, XElement Element);
}
