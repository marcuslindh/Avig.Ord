using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Avig.Ord.Internal;

namespace Avig.Ord;

/// <summary>
/// The formatting of text being written, set fluently: <c>s =&gt; s.Bold().Color("C00000")</c>.
/// Each setting is written on the run itself, as Word writes direct formatting.
/// </summary>
public sealed class RunStyle
{
    private bool? bold;
    private bool? italic;
    private bool? underline;
    private bool? strike;
    private Length? size;
    private string? font;
    private string? color;
    private string? highlight;
    private VerticalPosition? position;
    private bool? allCaps;
    private bool? smallCaps;
    private bool? hidden;
    private string? characterStyle;

    /// <summary>Bold, or not.</summary>
    public RunStyle Bold(bool on = true)
    {
        bold = on;

        return this;
    }

    /// <summary>Italic, or not.</summary>
    public RunStyle Italic(bool on = true)
    {
        italic = on;

        return this;
    }

    /// <summary>A single underline, or none.</summary>
    public RunStyle Underline(bool on = true)
    {
        underline = on;

        return this;
    }

    /// <summary>Struck through, or not.</summary>
    public RunStyle Strike(bool on = true)
    {
        strike = on;

        return this;
    }

    /// <summary>The font size. Word keeps sizes in half-points, so 10.25pt becomes 10.5pt.</summary>
    public RunStyle Size(Length fontSize)
    {
        size = fontSize;

        return this;
    }

    /// <summary>The font size in points.</summary>
    public RunStyle Size(double points)
    {
        return Size(Length.Pt(points));
    }

    /// <summary>The font, such as "Arial".</summary>
    public RunStyle Font(string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        font = name;

        return this;
    }

    /// <summary>The text colour as RRGGBB, such as "C00000" (a leading # is allowed).</summary>
    /// <exception cref="ArgumentException">Not six hexadecimal digits.</exception>
    public RunStyle Color(string rgb)
    {
        color = Hex(rgb, nameof(rgb));

        return this;
    }

    /// <summary>A highlight by Word's name for it: yellow, green, cyan, magenta, blue, red, darkBlue, darkCyan, darkGreen, darkMagenta, darkRed, darkYellow, darkGray, lightGray, black.</summary>
    public RunStyle Highlight(string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        highlight = name;

        return this;
    }

    /// <summary>Raised and smaller, as in x².</summary>
    public RunStyle Superscript()
    {
        position = VerticalPosition.Superscript;

        return this;
    }

    /// <summary>Lowered and smaller, as in H₂O.</summary>
    public RunStyle Subscript()
    {
        position = VerticalPosition.Subscript;

        return this;
    }

    /// <summary>Shown in capitals.</summary>
    public RunStyle AllCaps(bool on = true)
    {
        allCaps = on;

        return this;
    }

    /// <summary>Shown in small capitals.</summary>
    public RunStyle SmallCaps(bool on = true)
    {
        smallCaps = on;

        return this;
    }

    /// <summary>Hidden text.</summary>
    public RunStyle Hidden(bool on = true)
    {
        hidden = on;

        return this;
    }

    /// <summary>A character style by id, such as Strong or Emphasis. The style must be in the document, or be one of those the package knows (Strong, Emphasis, Hyperlink).</summary>
    public RunStyle CharacterStyle(string styleId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleId);
        characterStyle = styleId;

        return this;
    }

    internal static string Hex(string rgb, string parameter)
    {
        string value = rgb.TrimStart('#');

        if (value.Length != 6 || value.All(char.IsAsciiHexDigit) == false)
        {
            throw new ArgumentException($"'{rgb}' is not a colour as RRGGBB, such as C00000.", parameter);
        }

        return value.ToUpperInvariant();
    }

    /// <summary>w:rPr with the settings in schema order; null when nothing is set.</summary>
    internal XElement? ToProperties(IContentSink sink)
    {
        WordNames names = sink.Names;
        XNamespace wml = names.W;
        XElement properties = new(wml + "rPr");

        if (characterStyle is not null)
        {
            sink.UseStyle(characterStyle);
            properties.Add(Value(names, "rStyle", characterStyle));
        }

        if (font is not null)
        {
            properties.Add(new XElement(wml + "rFonts", new XAttribute(wml + "ascii", font), new XAttribute(wml + "hAnsi", font), new XAttribute(wml + "cs", font)));
        }

        Toggle(properties, names, "b", bold);
        Toggle(properties, names, "bCs", bold);
        Toggle(properties, names, "i", italic);
        Toggle(properties, names, "iCs", italic);
        Toggle(properties, names, "caps", allCaps);
        Toggle(properties, names, "smallCaps", smallCaps);
        Toggle(properties, names, "strike", strike);
        Toggle(properties, names, "vanish", hidden);

        if (color is not null)
        {
            properties.Add(Value(names, "color", color));
        }

        if (size is not null)
        {
            string halfPoints = Math.Max(size.Value.ToHalfPoints(), 1).ToString(CultureInfo.InvariantCulture);
            properties.Add(Value(names, "sz", halfPoints), Value(names, "szCs", halfPoints));
        }

        if (highlight is not null)
        {
            properties.Add(Value(names, "highlight", highlight));
        }

        if (underline is not null)
        {
            string kind = "none";

            if (underline == true)
            {
                kind = "single";
            }

            properties.Add(Value(names, "u", kind));
        }

        if (position is not null)
        {
            string value = position switch
            {
                VerticalPosition.Superscript => "superscript",
                VerticalPosition.Subscript => "subscript",
                _ => "baseline",
            };

            properties.Add(Value(names, "vertAlign", value));
        }

        XElement? result = null;

        if (properties.HasElements == true)
        {
            result = properties;
        }

        return result;
    }

    internal static XElement Value(WordNames names, string name, string value)
    {
        return new XElement(names.W + name, new XAttribute(names.Val, value));
    }

    private static void Toggle(XElement properties, WordNames names, string name, bool? state)
    {
        if (state == true)
        {
            properties.Add(new XElement(names.W + name));
        }
        else if (state == false)
        {
            properties.Add(Value(names, name, "0"));
        }
    }
}

/// <summary>
/// Builds a paragraph: its text, links, pictures, notes and fields, and its style, alignment,
/// spacing and list.
/// </summary>
/// <example>
/// <code>
/// await writer.ParagraphAsync(p => p
///     .Text("Total: ")
///     .Text("1 200 kr", style => style.Bold())
///     .Link(" (details)", "https://example.com/"));
/// </code>
/// </example>
public sealed class ParagraphBuilder
{
    private readonly List<Func<IContentSink, IEnumerable<XElement>>> content = [];
    private string? style;
    private Alignment? alignment;
    private Length? spaceBefore;
    private Length? spaceAfter;
    private Length? indentStart;
    private Length? indentEnd;
    private Length? firstLine;
    private bool keepWithNext;
    private bool pageBreakBefore;
    private (bool Bullet, int Level, bool StartNew)? list;
    private RunStyle? paragraphMarkStyle;

    /// <summary>Text, with tabs ('\t') and line breaks ('\n') as Word's own, in an optional style.</summary>
    public ParagraphBuilder Text(string text, Action<RunStyle>? style = null)
    {
        ArgumentNullException.ThrowIfNull(text);
        RunStyle? runStyle = Style(style);
        content.Add(sink => [ContentXml.Run(sink, text, runStyle)]);

        return this;
    }

    /// <summary>A hyperlink to a URL, styled as Word styles links (blue, underlined) unless <paramref name="style"/> says otherwise.</summary>
    /// <exception cref="ArgumentException">The URL is not an absolute URL.</exception>
    public ParagraphBuilder Link(string text, string url, Action<RunStyle>? style = null)
    {
        ArgumentNullException.ThrowIfNull(text);

        if (Uri.TryCreate(url, UriKind.Absolute, out Uri? uri) == false)
        {
            throw new ArgumentException($"'{url}' is not an absolute URL.", nameof(url));
        }

        RunStyle runStyle = new();
        runStyle.CharacterStyle("Hyperlink");
        style?.Invoke(runStyle);
        string target = uri.AbsoluteUri;

        content.Add(sink =>
        {
            string id = sink.AddHyperlink(target);
            XElement link = new(sink.Names.Hyperlink, new XAttribute(sink.Names.Id, id), new XAttribute(sink.Names.W + "history", "1"), ContentXml.Run(sink, text, runStyle));

            return [link];
        });

        return this;
    }

    /// <summary>A tab.</summary>
    public ParagraphBuilder Tab()
    {
        return Text("\t");
    }

    /// <summary>A line break: the text goes on on the next line, in the same paragraph.</summary>
    public ParagraphBuilder LineBreak()
    {
        return Text("\n");
    }

    /// <summary>A page break: the text goes on on the next page.</summary>
    public ParagraphBuilder PageBreak()
    {
        content.Add(sink => [new XElement(sink.Names.Run, new XElement(sink.Names.W + "br", new XAttribute(sink.Names.Type, "page")))]);

        return this;
    }

    /// <summary>
    /// A picture in the line of text, at its own size at 96 pixels per inch (no wider than the
    /// text). PNG, JPEG, GIF and BMP; an SVG needs a PNG to fall back on (see
    /// <see cref="Picture(byte[], byte[], Length, Length, string)"/>).
    /// </summary>
    /// <exception cref="ArgumentException">The bytes are not a picture of a format Word shows, or its size cannot be read.</exception>
    public ParagraphBuilder Picture(byte[] data, string description = "")
    {
        ArgumentNullException.ThrowIfNull(data);

        if (ImageInfo.TryGetSize(data, out int width, out int height) == false)
        {
            throw new ArgumentException("The size of the picture cannot be read from it. Give its width and height.", nameof(data));
        }

        Length pictureWidth = Length.Px(width);
        Length pictureHeight = Length.Px(height);

        content.Add(sink =>
        {
            Length shownWidth = pictureWidth;
            Length shownHeight = pictureHeight;

            if (shownWidth > sink.TextWidth && sink.TextWidth > Length.Zero)
            {
                shownHeight = Length.FromEmu((long)Math.Round((double)shownHeight.Emu * sink.TextWidth.Emu / shownWidth.Emu));
                shownWidth = sink.TextWidth;
            }

            return [ContentXml.Picture(sink, data, null, shownWidth, shownHeight, description)];
        });

        return this;
    }

    /// <summary>A picture at a given size.</summary>
    /// <exception cref="ArgumentException">The bytes are not a PNG, JPEG, GIF or BMP.</exception>
    public ParagraphBuilder Picture(byte[] data, Length width, Length height, string description = "")
    {
        ArgumentNullException.ThrowIfNull(data);
        content.Add(sink => [ContentXml.Picture(sink, data, null, width, height, description)]);

        return this;
    }

    /// <summary>An SVG picture, with the PNG that Word 2013 and other programs that cannot show SVG use instead.</summary>
    /// <exception cref="ArgumentException">The first is not an SVG or the second not a PNG.</exception>
    public ParagraphBuilder Picture(byte[] svg, byte[] fallbackPng, Length width, Length height, string description = "")
    {
        ArgumentNullException.ThrowIfNull(svg);
        ArgumentNullException.ThrowIfNull(fallbackPng);
        content.Add(sink => [ContentXml.Picture(sink, fallbackPng, svg, width, height, description)]);

        return this;
    }

    /// <summary>A footnote: the mark here, and the text at the bottom of the page.</summary>
    public ParagraphBuilder Footnote(string text)
    {
        return Note(NoteKind.Footnote, text);
    }

    /// <summary>An endnote: the mark here, and the text at the end of the document.</summary>
    public ParagraphBuilder Endnote(string text)
    {
        return Note(NoteKind.Endnote, text);
    }

    /// <summary>The page number, as a PAGE field that Word fills in.</summary>
    public ParagraphBuilder PageNumber(Action<RunStyle>? style = null)
    {
        return Field("PAGE", "1", style);
    }

    /// <summary>The number of pages, as a NUMPAGES field that Word fills in.</summary>
    public ParagraphBuilder PageCount(Action<RunStyle>? style = null)
    {
        return Field("NUMPAGES", "1", style);
    }

    /// <summary>A MERGEFIELD, for Word's mail merge and for the templates of OrdDocument.FillTemplate.</summary>
    public ParagraphBuilder MergeField(string name, Action<RunStyle>? style = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        string quoted = name;

        if (name.Contains(' ', StringComparison.Ordinal) == true)
        {
            quoted = "\"" + name + "\"";
        }

        return Field("MERGEFIELD " + quoted + " \\* MERGEFORMAT", "«" + name + "»", style);
    }

    /// <summary>A field with an instruction, such as <c>DATE \@ "yyyy-MM-dd"</c>, and the result it shows until Word updates it.</summary>
    public ParagraphBuilder Field(string instruction, string result, Action<RunStyle>? style = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(instruction);
        ArgumentNullException.ThrowIfNull(result);
        RunStyle? runStyle = Style(style);
        content.Add(sink => ContentXml.Field(sink, instruction, result, runStyle));

        return this;
    }

    /// <summary>A paragraph style by id, such as Quote or one of the document's own.</summary>
    public ParagraphBuilder Style(string styleId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleId);
        style = styleId;

        return this;
    }

    /// <summary>A heading, level 1 to 9, in the style Heading1 … Heading9.</summary>
    public ParagraphBuilder Heading(int level)
    {
        if (level is < 1 or > 9)
        {
            throw new ArgumentOutOfRangeException(nameof(level), level, "A heading level is 1 to 9.");
        }

        style = BuiltInStyles.Heading(level);

        return this;
    }

    /// <summary>How the lines line up.</summary>
    public ParagraphBuilder Align(Alignment value)
    {
        alignment = value;

        return this;
    }

    /// <summary>Space above and below the paragraph.</summary>
    public ParagraphBuilder Spacing(Length? before = null, Length? after = null)
    {
        spaceBefore = before;
        spaceAfter = after;

        return this;
    }

    /// <summary>Indents from the margins; a negative <paramref name="firstLine"/> is a hanging indent.</summary>
    public ParagraphBuilder Indent(Length? start = null, Length? end = null, Length? firstLine = null)
    {
        indentStart = start;
        indentEnd = end;
        this.firstLine = firstLine;

        return this;
    }

    /// <summary>Kept on the same page as the next paragraph.</summary>
    public ParagraphBuilder KeepWithNext()
    {
        keepWithNext = true;

        return this;
    }

    /// <summary>Starts on a new page.</summary>
    public ParagraphBuilder PageBreakBefore()
    {
        pageBreakBefore = true;

        return this;
    }

    /// <summary>An item of a bulleted list, at a level from 0. Consecutive bulleted paragraphs make one list.</summary>
    public ParagraphBuilder Bullet(int level = 0)
    {
        list = (true, Level(level), false);

        return this;
    }

    /// <summary>
    /// An item of a numbered list, at a level from 0. The item goes on numbering the last
    /// numbered list; <paramref name="startNew"/> starts a new one at 1.
    /// </summary>
    public ParagraphBuilder Numbered(int level = 0, bool startNew = false)
    {
        list = (false, Level(level), startNew);

        return this;
    }

    /// <summary>The formatting of the paragraph mark, which Word uses for an empty paragraph's height and the list label.</summary>
    public ParagraphBuilder MarkStyle(Action<RunStyle> style)
    {
        ArgumentNullException.ThrowIfNull(style);
        paragraphMarkStyle = new RunStyle();
        style(paragraphMarkStyle);

        return this;
    }

    internal XElement Build(IContentSink sink)
    {
        WordNames names = sink.Names;
        XNamespace wml = names.W;
        XElement paragraph = new(names.Paragraph);
        XElement properties = new(names.ParagraphProperties);
        string? styleId = style;

        if (list is not null && styleId is null)
        {
            styleId = "ListParagraph";
        }

        if (styleId is not null)
        {
            sink.UseStyle(styleId);
            properties.Add(RunStyle.Value(names, "pStyle", styleId));
        }

        if (keepWithNext == true)
        {
            properties.Add(new XElement(wml + "keepNext"));
        }

        if (pageBreakBefore == true)
        {
            properties.Add(new XElement(wml + "pageBreakBefore"));
        }

        if (list is not null)
        {
            string numberingId = sink.ListNumbering(list.Value.Bullet, list.Value.StartNew);
            properties.Add(new XElement(wml + "numPr", RunStyle.Value(names, "ilvl", list.Value.Level.ToString(CultureInfo.InvariantCulture)), RunStyle.Value(names, "numId", numberingId)));
        }

        if (spaceBefore is not null || spaceAfter is not null)
        {
            XElement spacing = new(wml + "spacing");

            if (spaceBefore is not null)
            {
                spacing.Add(new XAttribute(wml + "before", Twips(spaceBefore.Value)));
            }

            if (spaceAfter is not null)
            {
                spacing.Add(new XAttribute(wml + "after", Twips(spaceAfter.Value)));
            }

            properties.Add(spacing);
        }

        if (indentStart is not null || indentEnd is not null || firstLine is not null)
        {
            XElement indent = new(wml + "ind");

            if (indentStart is not null)
            {
                indent.Add(new XAttribute(wml + "left", Twips(indentStart.Value)));
            }

            if (indentEnd is not null)
            {
                indent.Add(new XAttribute(wml + "right", Twips(indentEnd.Value)));
            }

            if (firstLine is not null && firstLine.Value < Length.Zero)
            {
                indent.Add(new XAttribute(wml + "hanging", Twips(Length.Zero - firstLine.Value)));
            }
            else if (firstLine is not null)
            {
                indent.Add(new XAttribute(wml + "firstLine", Twips(firstLine.Value)));
            }

            properties.Add(indent);
        }

        if (alignment is not null)
        {
            string value = alignment switch
            {
                Alignment.Center => "center",
                Alignment.End => "right",
                Alignment.Justified => "both",
                _ => "left",
            };

            properties.Add(RunStyle.Value(names, "jc", value));
        }

        XElement? mark = paragraphMarkStyle?.ToProperties(sink);

        if (mark is not null)
        {
            properties.Add(mark);
        }

        if (properties.HasElements == true)
        {
            paragraph.Add(properties);
        }

        foreach (Func<IContentSink, IEnumerable<XElement>> part in content)
        {
            paragraph.Add(part(sink));
        }

        return paragraph;
    }

    private static RunStyle? Style(Action<RunStyle>? style)
    {
        RunStyle? runStyle = null;

        if (style is not null)
        {
            runStyle = new RunStyle();
            style(runStyle);
        }

        return runStyle;
    }

    private static int Level(int level)
    {
        if (level is < 0 or > 8)
        {
            throw new ArgumentOutOfRangeException(nameof(level), level, "A list level is 0 to 8.");
        }

        return level;
    }

    private static string Twips(Length length)
    {
        return length.ToTwips().ToString(CultureInfo.InvariantCulture);
    }

    private ParagraphBuilder Note(NoteKind kind, string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        content.Add(sink =>
        {
            WordNames names = sink.Names;
            string textStyle = "FootnoteText";
            string referenceStyle = "FootnoteReference";
            string mark = "footnoteRef";
            string reference = "footnoteReference";

            if (kind == NoteKind.Endnote)
            {
                textStyle = "EndnoteText";
                referenceStyle = "EndnoteReference";
                mark = "endnoteRef";
                reference = "endnoteReference";
            }

            sink.UseStyle(textStyle);
            sink.UseStyle(referenceStyle);
            XElement referenceProperties = new(names.RunProperties, RunStyle.Value(names, "rStyle", referenceStyle));
            XElement note = new(
                names.Paragraph,
                new XElement(names.ParagraphProperties, RunStyle.Value(names, "pStyle", textStyle)),
                new XElement(names.Run, new XElement(referenceProperties), new XElement(names.W + mark)),
                ContentXml.Run(sink, " " + text, null));
            string id = sink.AddNote(kind, [note]);

            return [new XElement(names.Run, new XElement(referenceProperties), new XElement(names.W + reference, new XAttribute(names.WId, id)))];
        });

        return this;
    }
}

/// <summary>Builds a table: its rows, cells, column widths and style.</summary>
/// <example>
/// <code>
/// await writer.TableAsync(t => t
///     .HeaderRow("Name", "Amount")
///     .Row("Rent", "8 000")
///     .Row("Food", "3 500"));
/// </code>
/// </example>
public sealed class TableBuilder
{
    private readonly List<RowBuilder> rows = [];
    private string style = "TableGrid";
    private Length[]? widths;

    /// <summary>A header row, which Word repeats at the top of every page the table runs over; its text is bold.</summary>
    public TableBuilder HeaderRow(params string[] cells)
    {
        ArgumentNullException.ThrowIfNull(cells);

        return Row(row =>
        {
            row.Header();

            foreach (string cell in cells)
            {
                row.Cell(cell, built => built.Paragraph(paragraph => paragraph.Text(cell, style => style.Bold())));
            }
        });
    }

    /// <summary>A row of cells with text.</summary>
    public TableBuilder Row(params string[] cells)
    {
        ArgumentNullException.ThrowIfNull(cells);

        return Row(row =>
        {
            foreach (string cell in cells)
            {
                row.Cell(cell);
            }
        });
    }

    /// <summary>A row built cell by cell.</summary>
    public TableBuilder Row(Action<RowBuilder> build)
    {
        ArgumentNullException.ThrowIfNull(build);
        RowBuilder row = new();
        build(row);
        rows.Add(row);

        return this;
    }

    /// <summary>The table style by id; TableGrid (a thin border round every cell) unless set.</summary>
    public TableBuilder Style(string styleId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleId);
        style = styleId;

        return this;
    }

    /// <summary>The columns' widths. Without them the columns share the width between the margins.</summary>
    public TableBuilder ColumnWidths(params Length[] columnWidths)
    {
        ArgumentNullException.ThrowIfNull(columnWidths);
        widths = columnWidths;

        return this;
    }

    internal XElement Build(IContentSink sink, Length available)
    {
        WordNames names = sink.Names;
        XNamespace wml = names.W;
        int columns = Math.Max(1, rows.Select(row => row.Columns).DefaultIfEmpty(1).Max());

        if (widths is not null)
        {
            columns = Math.Max(columns, widths.Length);
        }

        Length[] grid = new Length[columns];

        for (int column = 0; column < columns; column++)
        {
            if (widths is not null && column < widths.Length)
            {
                grid[column] = widths[column];
            }
            else
            {
                grid[column] = Length.FromEmu(available.Emu / columns);
            }
        }

        sink.UseStyle(style);
        XElement table = new(
            names.Table,
            new XElement(
                names.TableProperties,
                RunStyle.Value(names, "tblStyle", style),
                new XElement(wml + "tblW", new XAttribute(wml + "w", "0"), new XAttribute(wml + "type", "auto")),
                new XElement(wml + "tblLook", new XAttribute(wml + "val", "04A0"), new XAttribute(wml + "firstRow", "1"), new XAttribute(wml + "lastRow", "0"), new XAttribute(wml + "firstColumn", "1"), new XAttribute(wml + "lastColumn", "0"), new XAttribute(wml + "noHBand", "0"), new XAttribute(wml + "noVBand", "1"))),
            new XElement(names.TableGrid, grid.Select(width => new XElement(wml + "gridCol", new XAttribute(wml + "w", width.ToTwips().ToString(CultureInfo.InvariantCulture))))));

        if (rows.Count == 0)
        {
            RowBuilder empty = new();
            empty.Cell(string.Empty);
            rows.Add(empty);
        }

        foreach (RowBuilder row in rows)
        {
            table.Add(row.Build(sink, grid));
        }

        return table;
    }
}

/// <summary>Builds a table row.</summary>
public sealed class RowBuilder
{
    private readonly List<CellBuilder> cells = [];
    private bool header;

    internal int Columns => cells.Sum(cell => cell.Columns);

    /// <summary>Makes the row a header row, repeated at the top of every page.</summary>
    public RowBuilder Header()
    {
        header = true;

        return this;
    }

    /// <summary>A cell with text, and optionally more: <c>row.Cell("x", c =&gt; c.Span(2))</c>.</summary>
    public RowBuilder Cell(string text, Action<CellBuilder>? build = null)
    {
        ArgumentNullException.ThrowIfNull(text);
        CellBuilder cell = new();

        if (build is null)
        {
            cell.Paragraph(text);
        }
        else
        {
            build(cell);

            if (cell.HasContent == false)
            {
                cell.Paragraph(text);
            }
        }

        cells.Add(cell);

        return this;
    }

    /// <summary>A cell built with paragraphs and tables of its own.</summary>
    public RowBuilder Cell(Action<CellBuilder> build)
    {
        ArgumentNullException.ThrowIfNull(build);
        CellBuilder cell = new();
        build(cell);
        cells.Add(cell);

        return this;
    }

    internal XElement Build(IContentSink sink, Length[] grid)
    {
        WordNames names = sink.Names;
        XElement row = new(names.TableRow);

        if (header == true)
        {
            row.Add(new XElement(names.TableRowProperties, new XElement(names.W + "tblHeader")));
        }

        int column = 0;

        foreach (CellBuilder cell in cells)
        {
            int span = Math.Min(cell.Columns, Math.Max(grid.Length - column, 1));
            Length width = Length.Zero;

            for (int index = column; index < column + span && index < grid.Length; index++)
            {
                width += grid[index];
            }

            row.Add(cell.Build(sink, width, span));
            column += span;
        }

        // A row with fewer cells than the grid gets empty ones, as Word would show them.
        while (column < grid.Length)
        {
            CellBuilder empty = new();
            row.Add(empty.Build(sink, grid[column], 1));
            column++;
        }

        return row;
    }
}

/// <summary>Builds a table cell: its paragraphs and tables, span, merge, width and shading.</summary>
public sealed class CellBuilder
{
    private readonly List<Func<IContentSink, Length, XElement>> blocks = [];
    private string? shading;
    private VerticalMerge merge;

    internal int Columns { get; private set; } = 1;

    internal bool HasContent => blocks.Count > 0;

    /// <summary>A paragraph of plain text.</summary>
    public CellBuilder Paragraph(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        return Paragraph(paragraph => paragraph.Text(text));
    }

    /// <summary>A paragraph built with <see cref="ParagraphBuilder"/>.</summary>
    public CellBuilder Paragraph(Action<ParagraphBuilder> build)
    {
        ArgumentNullException.ThrowIfNull(build);
        ParagraphBuilder paragraph = new();
        build(paragraph);
        blocks.Add((sink, _) => paragraph.Build(sink));

        return this;
    }

    /// <summary>A table inside the cell.</summary>
    public CellBuilder Table(Action<TableBuilder> build)
    {
        ArgumentNullException.ThrowIfNull(build);
        TableBuilder table = new();
        build(table);
        blocks.Add((sink, width) => table.Build(sink, width - Length.Pt(10.8)));

        return this;
    }

    /// <summary>How many grid columns the cell spans.</summary>
    public CellBuilder Span(int columns)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(columns, 1);
        Columns = columns;

        return this;
    }

    /// <summary>Starts a vertical merge: the cells below that <see cref="MergeUp"/> join this one.</summary>
    public CellBuilder MergeDown()
    {
        merge = VerticalMerge.Start;

        return this;
    }

    /// <summary>Joins the cell above, which has <see cref="MergeDown"/> or this; the cell's own content is not shown.</summary>
    public CellBuilder MergeUp()
    {
        merge = VerticalMerge.Continue;

        return this;
    }

    /// <summary>A background colour as RRGGBB.</summary>
    public CellBuilder Shading(string rgb)
    {
        shading = RunStyle.Hex(rgb, nameof(rgb));

        return this;
    }

    internal XElement Build(IContentSink sink, Length width, int span)
    {
        WordNames names = sink.Names;
        XNamespace wml = names.W;
        XElement properties = new(
            names.TableCellProperties,
            new XElement(wml + "tcW", new XAttribute(wml + "w", width.ToTwips().ToString(CultureInfo.InvariantCulture)), new XAttribute(wml + "type", "dxa")));

        if (span > 1)
        {
            properties.Add(RunStyle.Value(names, "gridSpan", span.ToString(CultureInfo.InvariantCulture)));
        }

        if (merge == VerticalMerge.Start)
        {
            properties.Add(RunStyle.Value(names, "vMerge", "restart"));
        }
        else if (merge == VerticalMerge.Continue)
        {
            properties.Add(new XElement(wml + "vMerge"));
        }

        if (shading is not null)
        {
            properties.Add(new XElement(wml + "shd", new XAttribute(names.Val, "clear"), new XAttribute(wml + "color", "auto"), new XAttribute(wml + "fill", shading)));
        }

        XElement cell = new(names.TableCell, properties);

        foreach (Func<IContentSink, Length, XElement> block in blocks)
        {
            cell.Add(block(sink, width));
        }

        // A cell must end with a paragraph, also when it is empty or ends with a table.
        if (cell.Elements().LastOrDefault()?.Name != names.Paragraph)
        {
            cell.Add(new XElement(names.Paragraph));
        }

        return cell;
    }
}

/// <summary>Makes the XML of runs, fields and pictures for the builders.</summary>
internal static class ContentXml
{
    public static XElement Run(IContentSink sink, string text, RunStyle? style)
    {
        WordNames names = sink.Names;
        XElement run = new(names.Run);
        XElement? properties = style?.ToProperties(sink);

        if (properties is not null)
        {
            run.Add(properties);
        }

        AddText(names, run, text);

        return run;
    }

    /// <summary>Text as w:t, with '\t' as w:tab and '\n' as w:br; characters XML cannot hold are left out.</summary>
    public static void AddText(WordNames names, XElement run, string text)
    {
        StringBuilder segment = new();

        foreach (char character in Clean(text))
        {
            if (character is '\t' or '\n')
            {
                Flush(names, run, segment);

                if (character == '\t')
                {
                    run.Add(new XElement(names.W + "tab"));
                }
                else
                {
                    run.Add(new XElement(names.W + "br"));
                }

                continue;
            }

            if (character != '\r')
            {
                segment.Append(character);
            }
        }

        Flush(names, run, segment);
    }

    /// <summary>The text without the characters XML 1.0 cannot hold (control characters but tab and newline, lone surrogates).</summary>
    public static string Clean(string text)
    {
        bool clean = true;

        foreach (char character in text)
        {
            if (XmlConvert.IsXmlChar(character) == false && char.IsSurrogate(character) == false)
            {
                clean = false;

                break;
            }
        }

        if (clean == true && HasLoneSurrogate(text) == false)
        {
            return text;
        }

        StringBuilder cleaned = new(text.Length);

        for (int index = 0; index < text.Length; index++)
        {
            char character = text[index];

            if (char.IsHighSurrogate(character) == true && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1]) == true)
            {
                cleaned.Append(character).Append(text[index + 1]);
                index++;

                continue;
            }

            if (XmlConvert.IsXmlChar(character) == true)
            {
                cleaned.Append(character);
            }
        }

        return cleaned.ToString();
    }

    public static IEnumerable<XElement> Field(IContentSink sink, string instruction, string result, RunStyle? style)
    {
        WordNames names = sink.Names;
        XElement? properties = style?.ToProperties(sink);
        XElement charRun(string type)
        {
            XElement run = new(names.Run);

            if (properties is not null)
            {
                run.Add(new XElement(properties));
            }

            run.Add(new XElement(names.FieldChar, new XAttribute(names.W + "fldCharType", type)));

            return run;
        }

        XElement instructionRun = new(names.Run);

        if (properties is not null)
        {
            instructionRun.Add(new XElement(properties));
        }

        instructionRun.Add(new XElement(names.InstructionText, new XAttribute(XNamespace.Xml + "space", "preserve"), " " + Clean(instruction) + " "));

        return [charRun("begin"), instructionRun, charRun("separate"), Run(sink, result, style), charRun("end")];
    }

    public static XElement Picture(IContentSink sink, byte[] data, byte[]? svg, Length width, Length height, string description)
    {
        string contentType = WordTypes.ImageContentType(data, out string extension);

        if (contentType is not ("image/png" or "image/jpeg" or "image/gif" or "image/bmp"))
        {
            throw new ArgumentException("The picture is not a PNG, JPEG, GIF or BMP. An SVG is added with a PNG to fall back on.", nameof(data));
        }

        string? svgId = null;

        if (svg is not null)
        {
            if (WordTypes.ImageContentType(svg, out string svgExtension) != "image/svg+xml")
            {
                throw new ArgumentException("The SVG picture is not an SVG.", nameof(svg));
            }

            svgId = sink.AddPicture(svg, "image/svg+xml", svgExtension);
        }

        string id = sink.AddPicture(data, contentType, extension);
        int drawingId = sink.NextDrawingId();
        XNamespace wp = WordNames.Drawing;
        XNamespace drawing = WordNames.DrawingMain;
        XNamespace pic = WordNames.Picture;
        WordNames names = sink.Names;
        string cx = Math.Max(width.Emu, 1).ToString(CultureInfo.InvariantCulture);
        string cy = Math.Max(height.Emu, 1).ToString(CultureInfo.InvariantCulture);
        string name = "Picture " + drawingId.ToString(CultureInfo.InvariantCulture);
        XElement blip = new(drawing + "blip", new XAttribute(names.Embed, id));

        if (svgId is not null)
        {
            blip.Add(new XElement(
                drawing + "extLst",
                new XElement(
                    drawing + "ext",
                    new XAttribute("uri", "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"),
                    new XElement(XName.Get("svgBlip", WordNames.SvgBlip), new XAttribute(XNamespace.Xmlns + "asvg", WordNames.SvgBlip), new XAttribute(names.Embed, svgId)))));
        }

        XElement inline = new(
            wp + "inline",
            new XAttribute("distT", "0"),
            new XAttribute("distB", "0"),
            new XAttribute("distL", "0"),
            new XAttribute("distR", "0"),
            new XElement(wp + "extent", new XAttribute("cx", cx), new XAttribute("cy", cy)),
            new XElement(wp + "effectExtent", new XAttribute("l", "0"), new XAttribute("t", "0"), new XAttribute("r", "0"), new XAttribute("b", "0")),
            new XElement(wp + "docPr", new XAttribute("id", drawingId.ToString(CultureInfo.InvariantCulture)), new XAttribute("name", name), new XAttribute("descr", Clean(description))),
            new XElement(wp + "cNvGraphicFramePr", new XElement(drawing + "graphicFrameLocks", new XAttribute(XNamespace.Xmlns + "a", drawing.NamespaceName), new XAttribute("noChangeAspect", "1"))),
            new XElement(
                drawing + "graphic",
                new XAttribute(XNamespace.Xmlns + "a", drawing.NamespaceName),
                new XElement(
                    drawing + "graphicData",
                    new XAttribute("uri", WordNames.Picture),
                    new XElement(
                        pic + "pic",
                        new XAttribute(XNamespace.Xmlns + "pic", pic.NamespaceName),
                        new XElement(pic + "nvPicPr", new XElement(pic + "cNvPr", new XAttribute("id", "0"), new XAttribute("name", name)), new XElement(pic + "cNvPicPr")),
                        new XElement(pic + "blipFill", blip, new XElement(drawing + "stretch", new XElement(drawing + "fillRect"))),
                        new XElement(
                            pic + "spPr",
                            new XElement(drawing + "xfrm", new XElement(drawing + "off", new XAttribute("x", "0"), new XAttribute("y", "0")), new XElement(drawing + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy))),
                            new XElement(drawing + "prstGeom", new XAttribute("prst", "rect"), new XElement(drawing + "avLst")))))));

        return new XElement(names.Run, new XElement(names.W + "drawing", inline));
    }

    private static bool HasLoneSurrogate(string text)
    {
        for (int index = 0; index < text.Length; index++)
        {
            if (char.IsHighSurrogate(text[index]) == true && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1]) == true)
            {
                index++;

                continue;
            }

            if (char.IsSurrogate(text[index]) == true)
            {
                return true;
            }
        }

        return false;
    }

    private static void Flush(WordNames names, XElement run, StringBuilder segment)
    {
        if (segment.Length == 0)
        {
            return;
        }

        string value = segment.ToString();
        XElement text = new(names.Text, value);

        if (value[0] == ' ' || value[^1] == ' ' || value.Contains("  ", StringComparison.Ordinal) == true)
        {
            text.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
        }

        run.Add(text);
        segment.Clear();
    }
}
