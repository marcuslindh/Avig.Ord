using System.Text;
using System.Xml.Linq;
using Avig.Ord.Internal;

namespace Avig.Ord;

/// <summary>
/// Something inside a paragraph: a run of text, a hyperlink, a picture, a field, a page break,
/// a reference to a note, a text box or an inline content control.
/// </summary>
public abstract class Inline
{
    private protected Inline(DocumentContext context, XElement element)
    {
        Context = context;
        Element = element;
    }

    /// <summary>The text it contributes to its paragraph's text.</summary>
    public abstract string Text { get; }

    internal DocumentContext Context { get; }

    internal XElement Element { get; }
}

/// <summary>
/// A run of text with one formatting. Tabs are '\t' and line breaks '\n' in <see cref="Text"/>;
/// a page or column break is a <see cref="Break"/> of its own.
/// </summary>
public sealed class Run : Inline
{
    private readonly string text;
    private readonly RunFormat inherited;

    internal Run(DocumentContext context, XElement element, string text, RevisionKind revision, RunFormat inherited)
        : base(context, element)
    {
        this.text = text;
        this.inherited = inherited;
        Revision = revision;
    }

    /// <inheritdoc/>
    public override string Text => text;

    /// <summary>Whether the run is part of a tracked change, when the document is read with <see cref="RevisionView.Marked"/>.</summary>
    public RevisionKind Revision { get; }

    /// <summary>The id of the run's character style, if it has one.</summary>
    public string? StyleId => Context.Names.ValueOf(Element.Element(Context.Names.RunProperties), "rStyle");

    /// <summary>What is set on the run itself.</summary>
    public RunFormat DirectFormat => FormatReader.Run(Context.Names, Element.Element(Context.Names.RunProperties));

    /// <summary>
    /// The formatting the run has, as Word shows it: the document defaults, the paragraph's
    /// style, the run's character style and what is set on the run itself, each over the one
    /// before.
    /// </summary>
    public RunFormat Format => inherited.With(Context.Styles.CharacterStyle(StyleId)).With(DirectFormat);
}

/// <summary>What kind of break.</summary>
public enum BreakKind
{
    /// <summary>The text goes on on the next page.</summary>
    Page,

    /// <summary>The text goes on in the next column.</summary>
    Column,
}

/// <summary>A page or column break inside a paragraph.</summary>
public sealed class Break : Inline
{
    internal Break(DocumentContext context, XElement element, BreakKind kind)
        : base(context, element)
    {
        Kind = kind;
    }

    /// <summary>Page or column.</summary>
    public BreakKind Kind { get; }

    /// <summary>A line break, '\n', so that the text before and after does not run together.</summary>
    public override string Text => "\n";
}

/// <summary>A hyperlink: to a URL, or to a bookmark in the document.</summary>
public sealed class Hyperlink : Inline
{
    internal Hyperlink(DocumentContext context, XElement element, IReadOnlyList<Inline> inlines)
        : base(context, element)
    {
        Inlines = inlines;
    }

    /// <summary>What the link shows.</summary>
    public IReadOnlyList<Inline> Inlines { get; }

    /// <inheritdoc/>
    public override string Text => string.Concat(Inlines.Select(inline => inline.Text));

    /// <summary>The URL, for a link out of the document; null for a link to a bookmark.</summary>
    public string? Url
    {
        get
        {
            string? url = null;
            string? id = Element.Attribute(Context.Names.Id)?.Value;

            if (id is not null && Context.TryResolve(id, out string target, out bool external) == true && external == true)
            {
                url = target;
            }

            return url;
        }
    }

    /// <summary>The bookmark the link goes to, for a link inside the document.</summary>
    public string? Anchor => Element.Attribute(Context.Names.W + "anchor")?.Value;
}

/// <summary>
/// A field: a page number, a date, a merge field, a table of contents … <see cref="Text"/> is
/// its result as Word last saved it; the package does not calculate fields again.
/// </summary>
public sealed class Field : Inline
{
    internal Field(DocumentContext context, XElement element, string instruction, IReadOnlyList<Inline> result)
        : base(context, element)
    {
        Instruction = instruction.Trim();
        Result = result;
    }

    /// <summary>The field code, such as <c>MERGEFIELD Name \* MERGEFORMAT</c> or <c>PAGE</c>.</summary>
    public string Instruction { get; }

    /// <summary>The first word of the field code, upper case: MERGEFIELD, PAGE, DATE, TOC, HYPERLINK …</summary>
    public string Code
    {
        get
        {
            int space = Instruction.IndexOf(' ', StringComparison.Ordinal);
            string code = Instruction;

            if (space >= 0)
            {
                code = Instruction[..space];
            }

            return code.ToUpperInvariant();
        }
    }

    /// <summary>The result as last saved.</summary>
    public IReadOnlyList<Inline> Result { get; }

    /// <inheritdoc/>
    public override string Text => string.Concat(Result.Select(inline => inline.Text));

    /// <summary>For a MERGEFIELD, the name of the field it fills in; empty for any other field.</summary>
    public string MergeFieldName
    {
        get
        {
            if (Code != "MERGEFIELD")
            {
                return string.Empty;
            }

            return FieldInstruction.FirstArgument(Instruction);
        }
    }
}

/// <summary>A picture in the text. Its bytes are in another part of the package; <see cref="ReadAsync"/> reads them.</summary>
public sealed class Picture : Inline
{
    internal Picture(DocumentContext context, XElement element, string? relationshipId, string? svgRelationshipId, Length width, Length height, string description, string title)
        : base(context, element)
    {
        RelationshipId = relationshipId;
        SvgRelationshipId = svgRelationshipId;
        Width = width;
        Height = height;
        Description = description;
        Title = title;
    }

    /// <summary>The picture's width on the page.</summary>
    public Length Width { get; }

    /// <summary>The picture's height on the page.</summary>
    public Length Height { get; }

    /// <summary>The alternative text, for screen readers (wp:docPr/@descr).</summary>
    public string Description { get; }

    /// <summary>The title (wp:docPr/@title).</summary>
    public string Title { get; }

    /// <summary>No text: a picture does not add to the paragraph's text.</summary>
    public override string Text => string.Empty;

    /// <summary>The part the picture's bytes are in, such as word/media/image1.png; empty for a linked picture.</summary>
    public string PartName
    {
        get
        {
            string name = string.Empty;

            if (RelationshipId is not null && Context.TryResolve(RelationshipId, out string target, out bool external) == true && external == false)
            {
                name = target;
            }

            return name;
        }
    }

    /// <summary>The part of the SVG version, when the picture is an SVG with a PNG to fall back on; empty otherwise.</summary>
    public string SvgPartName
    {
        get
        {
            string name = string.Empty;

            if (SvgRelationshipId is not null && Context.TryResolve(SvgRelationshipId, out string target, out bool external) == true && external == false)
            {
                name = target;
            }

            return name;
        }
    }

    internal string? RelationshipId { get; }

    internal string? SvgRelationshipId { get; }

    /// <summary>The picture's bytes; NotFound when it is linked from outside the file or its part is missing.</summary>
    public async Task<Result<byte[]>> ReadAsync(CancellationToken cancellationToken = default)
    {
        string part = PartName;

        if (part.Length == 0)
        {
            return Result<byte[]>.NotFound("The picture is not in the file: it is linked from outside, or its relationship is missing.");
        }

        return await Context.ReadPartAsync(part, cancellationToken);
    }
}

/// <summary>Which kind of note.</summary>
public enum NoteKind
{
    /// <summary>A footnote, at the bottom of the page.</summary>
    Footnote,

    /// <summary>An endnote, at the end of the document or section.</summary>
    Endnote,
}

/// <summary>The mark in the text that refers to a footnote or endnote.</summary>
public sealed class NoteReference : Inline
{
    internal NoteReference(DocumentContext context, XElement element, NoteKind kind, string id)
        : base(context, element)
    {
        Kind = kind;
        Id = id;
    }

    /// <summary>Footnote or endnote.</summary>
    public NoteKind Kind { get; }

    /// <summary>The note's id, as in <see cref="OrdDocument.Footnotes"/>.</summary>
    public string Id { get; }

    /// <summary>No text: Word numbers the mark itself when it lays out the page.</summary>
    public override string Text => string.Empty;
}

/// <summary>A text box in the text, with paragraphs of its own.</summary>
public sealed class TextBox : Inline
{
    internal TextBox(DocumentContext context, XElement element, IReadOnlyList<Block> blocks)
        : base(context, element)
    {
        Blocks = blocks;
    }

    /// <summary>The paragraphs and tables in the box.</summary>
    public IReadOnlyList<Block> Blocks { get; }

    /// <summary>No text: a text box sits beside the paragraph's text, not in it. Its own text is in <see cref="Blocks"/>.</summary>
    public override string Text => string.Empty;
}

/// <summary>An inline content control (w:sdt): a field in a form or a template, with a tag and a title.</summary>
public sealed class ContentControl : Inline
{
    internal ContentControl(DocumentContext context, XElement element, IReadOnlyList<Inline> inlines)
        : base(context, element)
    {
        Inlines = inlines;
    }

    /// <summary>What the control holds.</summary>
    public IReadOnlyList<Inline> Inlines { get; }

    /// <inheritdoc/>
    public override string Text => string.Concat(Inlines.Select(inline => inline.Text));

    /// <summary>The control's tag (w:tag), which code finds it by.</summary>
    public string Tag => Context.Names.ValueOf(Element.Element(Context.Names.ContentControlProperties), "tag") ?? string.Empty;

    /// <summary>The control's title (w:alias), which Word shows.</summary>
    public string Title => Context.Names.ValueOf(Element.Element(Context.Names.ContentControlProperties), "alias") ?? string.Empty;
}

/// <summary>Parses field instructions.</summary>
internal static class FieldInstruction
{
    /// <summary>The first argument after the field code: the name in <c>MERGEFIELD "First name" \* MERGEFORMAT</c>, unquoted.</summary>
    public static string FirstArgument(string instruction)
    {
        string rest = instruction.Trim();
        int space = rest.IndexOf(' ', StringComparison.Ordinal);

        if (space < 0)
        {
            return string.Empty;
        }

        rest = rest[(space + 1)..].TrimStart();

        if (rest.StartsWith('"') == true)
        {
            int end = rest.IndexOf('"', 1);

            if (end > 0)
            {
                return rest[1..end];
            }

            return rest[1..];
        }

        StringBuilder name = new();

        foreach (char character in rest)
        {
            if (char.IsWhiteSpace(character) == true)
            {
                break;
            }

            name.Append(character);
        }

        return name.ToString();
    }
}
