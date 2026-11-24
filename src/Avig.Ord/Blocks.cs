using System.Xml.Linq;
using Avig.Ord.Internal;

namespace Avig.Ord;

/// <summary>A paragraph or a table: the things a document body, a table cell, a header or a note is made of.</summary>
public abstract partial class Block
{
    private protected Block(DocumentContext context, XElement element)
    {
        Context = context;
        Element = element;
    }

    /// <summary>The text of the block: a paragraph's text, or a table's cells separated by tabs and its rows by line breaks.</summary>
    public abstract string Text { get; }

    /// <summary>True when the block was read with <see cref="DocxReader"/> and cannot be changed.</summary>
    public bool IsReadOnly => Context.IsReadOnly;

    internal DocumentContext Context { get; }

    internal XElement Element { get; }

    /// <summary>The paragraphs and tables of a body, cell, header or note, with content controls and custom XML unwrapped.</summary>
    internal static List<Block> Read(DocumentContext context, XElement container)
    {
        List<Block> blocks = [];
        Collect(context, container, blocks, 0);

        return blocks;
    }

    private static void Collect(DocumentContext context, XElement container, List<Block> blocks, int depth)
    {
        WordNames names = context.Names;

        foreach (XElement child in container.Elements())
        {
            if (child.Name == names.Paragraph)
            {
                blocks.Add(new Paragraph(context, child));
            }
            else if (child.Name == names.Table)
            {
                blocks.Add(new Table(context, child));
            }
            else if (depth < 32 && Unwraps(context, child) == true)
            {
                XElement inner = child.Element(names.ContentControlContent) ?? child;
                Collect(context, inner, blocks, depth + 1);
            }
        }
    }

    /// <summary>Elements whose paragraphs count as the container's own: a block content control, custom XML, and an inserted or deleted block range as the view has it.</summary>
    private static bool Unwraps(DocumentContext context, XElement element)
    {
        WordNames names = context.Names;

        if (element.Name == names.ContentControl || element.Name == names.W + "customXml" || element.Name == names.W + "smartTag")
        {
            return true;
        }

        if (element.Name == names.Inserted || element.Name == names.MoveTo)
        {
            return context.View != RevisionView.Original;
        }

        if (element.Name == names.Deleted || element.Name == names.MoveFrom)
        {
            return context.View != RevisionView.Accepted;
        }

        return false;
    }
}

/// <summary>A paragraph: a line or more of text, a heading, a list item.</summary>
public sealed partial class Paragraph : Block
{
    internal Paragraph(DocumentContext context, XElement element)
        : base(context, element)
    {
    }

    /// <summary>The text as the document's <see cref="RevisionView"/> has it. Tabs are '\t' and line breaks '\n'; fields give their saved result.</summary>
    public override string Text => string.Concat(Inlines.Select(inline => inline.Text));

    /// <summary>What is in the paragraph, in order.</summary>
    public IReadOnlyList<Inline> Inlines => InlineReader.Read(Context, Element, InheritedRunFormat);

    /// <summary>Every run in the paragraph, also those inside hyperlinks, fields and content controls.</summary>
    public IReadOnlyList<Run> Runs => [.. InlineReader.Flatten(Inlines).OfType<Run>()];

    /// <summary>The id of the paragraph's style; null when it has none and so has the document's default (usually Normal).</summary>
    public string? StyleId => Context.Names.ValueOf(Properties, "pStyle");

    /// <summary>The name Word shows for the paragraph's style, such as "heading 1"; the default style's name when it has none.</summary>
    public string StyleName
    {
        get
        {
            string name = string.Empty;
            string? id = StyleId ?? Context.Styles.DefaultParagraphStyle;

            if (id is not null && Context.Styles.TryGet(id, out Style? style) == true)
            {
                name = style.Name;
            }

            return name;
        }
    }

    /// <summary>What is set on the paragraph itself.</summary>
    public ParagraphFormat DirectFormat => FormatReader.Paragraph(Context.Names, Properties);

    /// <summary>The formatting the paragraph has, as Word shows it: the document defaults, its style, and what is set on it.</summary>
    public ParagraphFormat Format => Context.Styles.ParagraphStyle(StyleId).Paragraph.With(DirectFormat);

    /// <summary>The heading level, 0 for the top (Heading 1); null for a paragraph that is not a heading.</summary>
    public int? OutlineLevel
    {
        get
        {
            int? level = Format.OutlineLevel;

            if (level is not (>= 0 and <= 8))
            {
                level = null;
            }

            return level;
        }
    }

    /// <summary>True for a heading: a paragraph with an outline level.</summary>
    public bool IsHeading => OutlineLevel is not null;

    /// <summary>
    /// The paragraph's place in a list, with the label Word shows; null when it is not a list
    /// item. Labels are counted through the whole document in order, so a paragraph from
    /// <see cref="DocxReader"/> has its label once the reader has come to it.
    /// </summary>
    public ListInfo? List
    {
        get
        {
            Context.Owner?.EnsureListLabels(Context);
            Context.TryGetListLabel(Element, out ListInfo? info);

            return info;
        }
    }

    /// <summary>True when a section ends with this paragraph: the page settings change after it.</summary>
    public bool EndsSection => Properties?.Element(Context.Names.SectionProperties) is not null;

    internal XElement? Properties => Element.Element(Context.Names.ParagraphProperties);

    internal RunFormat InheritedRunFormat => Context.Styles.ParagraphStyle(StyleId).Run;
}

/// <summary>A table.</summary>
public sealed partial class Table : Block
{
    internal Table(DocumentContext context, XElement element)
        : base(context, element)
    {
    }

    /// <summary>The rows, as the document's <see cref="RevisionView"/> has them.</summary>
    public IReadOnlyList<TableRow> Rows
    {
        get
        {
            List<TableRow> rows = [];

            foreach (XElement row in Element.Elements(Context.Names.TableRow))
            {
                if (TableRow.IsVisible(Context, row) == true)
                {
                    rows.Add(new TableRow(Context, row, this));
                }
            }

            return rows;
        }
    }

    /// <summary>How many columns the table's grid has.</summary>
    public int ColumnCount
    {
        get
        {
            int grid = Element.Element(Context.Names.TableGrid)?.Elements(Context.Names.W + "gridCol").Count() ?? 0;

            if (grid > 0)
            {
                return grid;
            }

            return Rows.Select(row => SpannedColumns(row)).DefaultIfEmpty(0).Max();
        }
    }

    /// <summary>The id of the table's style, such as TableGrid.</summary>
    public string? StyleId => Context.Names.ValueOf(Element.Element(Context.Names.TableProperties), "tblStyle");

    /// <summary>The cells separated by tabs and the rows by line breaks.</summary>
    public override string Text => string.Join('\n', Rows.Select(row => string.Join('\t', row.Cells.Select(cell => cell.Text))));

    /// <summary>The grid columns a row's cells span together, added up without overflow (Enumerable.Sum throws).</summary>
    internal static int SpannedColumns(TableRow row)
    {
        long columns = 0;

        foreach (TableCell cell in row.Cells)
        {
            columns += cell.ColumnSpan;
        }

        return (int)Math.Min(columns, int.MaxValue);
    }
}

/// <summary>A row of a table.</summary>
public sealed partial class TableRow
{
    internal TableRow(DocumentContext context, XElement element, Table table)
    {
        Context = context;
        Element = element;
        Table = table;
    }

    /// <summary>The table the row is in.</summary>
    public Table Table { get; }

    /// <summary>The cells, left to right.</summary>
    public IReadOnlyList<TableCell> Cells
    {
        get
        {
            List<TableCell> cells = [];
            Collect(Element, cells, 0);

            return cells;
        }
    }

    /// <summary>True for a header row, which Word repeats at the top of every page.</summary>
    public bool IsHeader => Context.Names.ToggleOf(Element.Element(Context.Names.TableRowProperties), "tblHeader") == true;

    /// <summary>The cells' text, separated by tabs.</summary>
    public string Text => string.Join('\t', Cells.Select(cell => cell.Text));

    internal DocumentContext Context { get; }

    internal XElement Element { get; }

    /// <summary>A row marked inserted or deleted is left out of the view that does not have it.</summary>
    internal static bool IsVisible(DocumentContext context, XElement row)
    {
        XElement? properties = row.Element(context.Names.TableRowProperties);

        if (properties?.Element(context.Names.Deleted) is not null && context.View == RevisionView.Accepted)
        {
            return false;
        }

        return properties?.Element(context.Names.Inserted) is null || context.View != RevisionView.Original;
    }

    private void Collect(XElement container, List<TableCell> cells, int depth)
    {
        foreach (XElement child in container.Elements())
        {
            if (child.Name == Context.Names.TableCell)
            {
                cells.Add(new TableCell(Context, child, this));
            }
            else if (depth < 8 && (child.Name == Context.Names.ContentControl || child.Name == Context.Names.W + "customXml"))
            {
                Collect(child.Element(Context.Names.ContentControlContent) ?? child, cells, depth + 1);
            }
        }
    }
}

/// <summary>How a cell takes part in a vertical merge.</summary>
public enum VerticalMerge
{
    /// <summary>Not merged with the cells above or below.</summary>
    None,

    /// <summary>The first cell of a merge: the merged cell's content is here.</summary>
    Start,

    /// <summary>Merged into the cell above; its own content is not shown.</summary>
    Continue,
}

/// <summary>A cell of a table, with paragraphs and tables of its own.</summary>
public sealed partial class TableCell
{
    internal TableCell(DocumentContext context, XElement element, TableRow row)
    {
        Context = context;
        Element = element;
        Row = row;
    }

    /// <summary>The row the cell is in.</summary>
    public TableRow Row { get; }

    /// <summary>The paragraphs and tables in the cell.</summary>
    public IReadOnlyList<Block> Blocks => Block.Read(Context, Element);

    /// <summary>The paragraphs directly in the cell (not those in tables inside it).</summary>
    public IReadOnlyList<Paragraph> Paragraphs => [.. Blocks.OfType<Paragraph>()];

    /// <summary>The cell's paragraphs' text, separated by line breaks.</summary>
    public string Text => string.Join('\n', Blocks.Select(block => block.Text));

    /// <summary>How many grid columns the cell spans; 1 unless it is merged across. At most 63, the most columns Word lets a table have.</summary>
    public int ColumnSpan => Math.Clamp(FormatReader.Integer(Context.Names.ValueOf(Properties, "gridSpan")) ?? 1, 1, MaxColumns);

    /// <summary>The most columns a table can have in Word. A larger gridSpan in a file overflowed a sum of spans, and made Markdown write that many empty cells.</summary>
    internal const int MaxColumns = 63;

    /// <summary>How the cell takes part in a vertical merge.</summary>
    public VerticalMerge VerticalMerge
    {
        get
        {
            XElement? merge = Properties?.Element(Context.Names.W + "vMerge");

            if (merge is null)
            {
                return VerticalMerge.None;
            }

            if (merge.Attribute(Context.Names.Val)?.Value == "restart")
            {
                return VerticalMerge.Start;
            }

            return VerticalMerge.Continue;
        }
    }

    internal DocumentContext Context { get; }

    internal XElement Element { get; }

    internal XElement? Properties => Element.Element(Context.Names.TableCellProperties);
}
