namespace Avig.Ord;

/// <summary>How a paragraph's lines line up between its indents.</summary>
public enum Alignment
{
    /// <summary>Against the start (left in a left-to-right language).</summary>
    Start,

    /// <summary>Centred.</summary>
    Center,

    /// <summary>Against the end (right in a left-to-right language).</summary>
    End,

    /// <summary>Stretched to both edges, except the last line.</summary>
    Justified,
}

/// <summary>Raised or lowered text.</summary>
public enum VerticalPosition
{
    /// <summary>On the line.</summary>
    Baseline,

    /// <summary>Raised and smaller, as in x².</summary>
    Superscript,

    /// <summary>Lowered and smaller, as in H₂O.</summary>
    Subscript,
}

/// <summary>
/// The formatting of a run of text. Every property is null when nothing in the chain sets it:
/// as read with <see cref="Run.Format"/> that means Word's own default (not bold, black, the
/// theme font); as <see cref="Run.DirectFormat"/>, that the run itself does not set it.
/// </summary>
public sealed record RunFormat
{
    /// <summary>No formatting at all.</summary>
    public static RunFormat None { get; } = new();

    /// <summary>Bold.</summary>
    public bool? Bold { get; init; }

    /// <summary>Italic.</summary>
    public bool? Italic { get; init; }

    /// <summary>Underlined (any underline style but none).</summary>
    public bool? Underline { get; init; }

    /// <summary>Struck through, once or twice.</summary>
    public bool? Strike { get; init; }

    /// <summary>The font size.</summary>
    public Length? Size { get; init; }

    /// <summary>The font for Latin text (w:rFonts/@w:ascii). Null when it comes from the theme.</summary>
    public string? Font { get; init; }

    /// <summary>The text colour as RRGGBB, or "auto".</summary>
    public string? Color { get; init; }

    /// <summary>The highlight, by Word's name for it: yellow, green, cyan …</summary>
    public string? Highlight { get; init; }

    /// <summary>Raised or lowered.</summary>
    public VerticalPosition? Position { get; init; }

    /// <summary>Shown in capitals.</summary>
    public bool? AllCaps { get; init; }

    /// <summary>Shown in small capitals.</summary>
    public bool? SmallCaps { get; init; }

    /// <summary>Hidden text, which Word does not show or print unless asked to.</summary>
    public bool? Hidden { get; init; }

    /// <summary>The properties of <paramref name="over"/> where it sets them, these where it does not.</summary>
    public RunFormat With(RunFormat over)
    {
        ArgumentNullException.ThrowIfNull(over);

        return new RunFormat
        {
            Bold = over.Bold ?? Bold,
            Italic = over.Italic ?? Italic,
            Underline = over.Underline ?? Underline,
            Strike = over.Strike ?? Strike,
            Size = over.Size ?? Size,
            Font = over.Font ?? Font,
            Color = over.Color ?? Color,
            Highlight = over.Highlight ?? Highlight,
            Position = over.Position ?? Position,
            AllCaps = over.AllCaps ?? AllCaps,
            SmallCaps = over.SmallCaps ?? SmallCaps,
            Hidden = over.Hidden ?? Hidden,
        };
    }
}

/// <summary>
/// The formatting of a paragraph. As with <see cref="RunFormat"/>, null means nothing in the
/// chain sets it.
/// </summary>
public sealed record ParagraphFormat
{
    /// <summary>No formatting at all.</summary>
    public static ParagraphFormat None { get; } = new();

    /// <summary>How the lines line up.</summary>
    public Alignment? Alignment { get; init; }

    /// <summary>Space above the paragraph.</summary>
    public Length? SpaceBefore { get; init; }

    /// <summary>Space below the paragraph.</summary>
    public Length? SpaceAfter { get; init; }

    /// <summary>The indent from the start margin.</summary>
    public Length? IndentStart { get; init; }

    /// <summary>The indent from the end margin.</summary>
    public Length? IndentEnd { get; init; }

    /// <summary>The first line's indent beyond <see cref="IndentStart"/>; negative for a hanging indent.</summary>
    public Length? FirstLineIndent { get; init; }

    /// <summary>
    /// The outline level, 0 for the top: what makes a paragraph a heading in the navigation
    /// pane, the table of contents and <see cref="OrdDocument.ToMarkdown(MarkdownOptions?)"/>.
    /// 9 is body text set explicitly (which overrides a heading style); null when nothing sets it.
    /// <see cref="Paragraph.OutlineLevel"/> gives the level with both meaning "not a heading".
    /// </summary>
    public int? OutlineLevel { get; init; }

    /// <summary>Kept on the same page as the next paragraph.</summary>
    public bool? KeepWithNext { get; init; }

    /// <summary>Starts on a new page.</summary>
    public bool? PageBreakBefore { get; init; }

    /// <summary>The properties of <paramref name="over"/> where it sets them, these where it does not.</summary>
    public ParagraphFormat With(ParagraphFormat over)
    {
        ArgumentNullException.ThrowIfNull(over);

        return new ParagraphFormat
        {
            Alignment = over.Alignment ?? Alignment,
            SpaceBefore = over.SpaceBefore ?? SpaceBefore,
            SpaceAfter = over.SpaceAfter ?? SpaceAfter,
            IndentStart = over.IndentStart ?? IndentStart,
            IndentEnd = over.IndentEnd ?? IndentEnd,
            FirstLineIndent = over.FirstLineIndent ?? FirstLineIndent,
            OutlineLevel = over.OutlineLevel ?? OutlineLevel,
            KeepWithNext = over.KeepWithNext ?? KeepWithNext,
            PageBreakBefore = over.PageBreakBefore ?? PageBreakBefore,
        };
    }
}
