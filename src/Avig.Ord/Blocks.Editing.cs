using System.Xml.Linq;
using Avig.Ord.Internal;

namespace Avig.Ord;

public abstract partial class Block
{
    /// <summary>Takes the block out of the document; while changes are tracked, marks it deleted instead.</summary>
    /// <exception cref="InvalidOperationException">The block was read with <see cref="DocxReader"/>.</exception>
    public void Remove()
    {
        Context.ThrowIfReadOnly();

        if (Context.Owner?.Tracker is Tracking tracking)
        {
            tracking.DeleteBlock(Context.Names, Element);
            Context.Changed();

            return;
        }

        XElement? parent = Element.Parent;
        Element.Remove();

        // A table cell must keep a paragraph.
        if (parent is not null && parent.Name == Context.Names.TableCell && parent.Elements(Context.Names.Paragraph).Any() == false)
        {
            parent.Add(new XElement(Context.Names.Paragraph));
        }

        Context.Changed();
    }

    /// <summary>Adds a paragraph after this block.</summary>
    /// <returns>The new paragraph.</returns>
    public Paragraph InsertParagraphAfter(Action<ParagraphBuilder> build)
    {
        XElement element = BuildParagraph(build);
        Element.AddAfterSelf(element);
        Context.Changed();

        return new Paragraph(Context, element);
    }

    /// <summary>Adds a paragraph of plain text after this block.</summary>
    /// <returns>The new paragraph.</returns>
    public Paragraph InsertParagraphAfter(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        return InsertParagraphAfter(paragraph => paragraph.Text(text));
    }

    /// <summary>Adds a paragraph before this block.</summary>
    /// <returns>The new paragraph.</returns>
    public Paragraph InsertParagraphBefore(Action<ParagraphBuilder> build)
    {
        XElement element = BuildParagraph(build);
        Element.AddBeforeSelf(element);
        Context.Changed();

        return new Paragraph(Context, element);
    }

    /// <summary>Adds a paragraph of plain text before this block.</summary>
    /// <returns>The new paragraph.</returns>
    public Paragraph InsertParagraphBefore(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        return InsertParagraphBefore(paragraph => paragraph.Text(text));
    }

    /// <summary>Adds a table after this block, with an empty paragraph after it so that it does not join a table that follows.</summary>
    /// <returns>The new table.</returns>
    public Table InsertTableAfter(Action<TableBuilder> build)
    {
        ArgumentNullException.ThrowIfNull(build);
        Context.ThrowIfReadOnly();
        TableBuilder builder = new();
        build(builder);
        IContentSink sink = Context.Owner!.SinkFor(Context);
        XElement element = builder.Build(sink, sink.TextWidth);
        Element.AddAfterSelf(element);
        XElement separator = new(Context.Names.Paragraph);
        element.AddAfterSelf(separator);
        Context.Owner.Tracker?.InsertBlock(Context.Names, element);
        Context.Owner.Tracker?.InsertBlock(Context.Names, separator);
        Context.Changed();

        return new Table(Context, element);
    }

    private XElement BuildParagraph(Action<ParagraphBuilder> build)
    {
        ArgumentNullException.ThrowIfNull(build);
        Context.ThrowIfReadOnly();
        ParagraphBuilder builder = new();
        build(builder);
        XElement element = builder.Build(Context.Owner!.SinkFor(Context));
        Context.Owner.Tracker?.InsertBlock(Context.Names, element);

        return element;
    }
}

public sealed partial class Paragraph
{
    /// <summary>Adds text at the end of the paragraph, in an optional style.</summary>
    /// <returns>The new run.</returns>
    /// <exception cref="InvalidOperationException">The paragraph was read with <see cref="DocxReader"/>.</exception>
    public Run AppendText(string text, Action<RunStyle>? style = null)
    {
        ArgumentNullException.ThrowIfNull(text);
        Context.ThrowIfReadOnly();
        RunStyle? runStyle = null;

        if (style is not null)
        {
            runStyle = new RunStyle();
            style(runStyle);
        }

        XElement run = ContentXml.Run(Context.Owner!.SinkFor(Context), text, runStyle);
        Element.Add(run);
        Context.Owner.Tracker?.Insert(Context.Names, [run]);
        Context.Changed();

        return new Run(Context, run, text, RevisionKind.None, InheritedRunFormat);
    }

    /// <summary>
    /// Replaces the paragraph's content with plain text in the formatting of its first run.
    /// Its style, alignment and list stay.
    /// </summary>
    /// <exception cref="InvalidOperationException">The paragraph was read with <see cref="DocxReader"/>.</exception>
    public void SetText(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        Context.ThrowIfReadOnly();
        WordNames names = Context.Names;
        XElement? firstProperties = Element.Descendants(names.Run).Select(run => run.Element(names.RunProperties)).FirstOrDefault(properties => properties is not null);
        Tracking? tracking = Context.Owner?.Tracker;

        if (tracking is not null)
        {
            tracking.Delete(names, [.. Element.Descendants(names.Run).Where(run => run.Ancestors(names.Paragraph).FirstOrDefault() == Element)]);
        }
        else
        {
            foreach (XNode node in Element.Nodes().ToList())
            {
                if (node is XElement element && element.Name == names.ParagraphProperties)
                {
                    continue;
                }

                node.Remove();
            }
        }

        XElement replacement = new(names.Run);

        if (firstProperties is not null)
        {
            replacement.Add(new XElement(firstProperties));
        }

        ContentXml.AddText(names, replacement, text);
        Element.Add(replacement);
        tracking?.Insert(names, [replacement]);
        Context.Changed();
    }

    /// <summary>Sets the paragraph's style by id, such as Heading2; null for the document's default style.</summary>
    /// <exception cref="InvalidOperationException">The paragraph was read with <see cref="DocxReader"/>.</exception>
    public void SetStyle(string? styleId)
    {
        Context.ThrowIfReadOnly();
        WordNames names = Context.Names;
        XElement? properties = Properties;

        if (styleId is null)
        {
            properties?.Element(names.W + "pStyle")?.Remove();
            Context.Changed();

            return;
        }

        Context.Owner!.SinkFor(Context).UseStyle(styleId);

        if (properties is null)
        {
            properties = new XElement(names.ParagraphProperties);
            Element.AddFirst(properties);
        }

        XElement? style = properties.Element(names.W + "pStyle");

        if (style is null)
        {
            properties.AddFirst(new XElement(names.W + "pStyle", new XAttribute(names.Val, styleId)));
        }
        else
        {
            style.SetAttributeValue(names.Val, styleId);
        }

        Context.Changed();
    }
}

public sealed partial class Table
{
    /// <summary>Adds a row of cells with text at the end of the table, formatted as the last row.</summary>
    /// <returns>The new row.</returns>
    /// <exception cref="InvalidOperationException">The table was read with <see cref="DocxReader"/>.</exception>
    public TableRow AddRow(params string[] cells)
    {
        ArgumentNullException.ThrowIfNull(cells);
        Context.ThrowIfReadOnly();
        IReadOnlyList<TableRow> rows = Rows;

        if (rows.Count == 0)
        {
            XElement row = new(Context.Names.TableRow);

            foreach (string text in cells)
            {
                XElement run = new(Context.Names.Run);
                ContentXml.AddText(Context.Names, run, text);
                row.Add(new XElement(Context.Names.TableCell, new XElement(Context.Names.Paragraph, run)));
            }

            Element.Add(row);
            Context.Changed();

            return new TableRow(Context, row, this);
        }

        TableRow copy = rows[^1].InsertCopyAfter();
        IReadOnlyList<TableCell> copied = copy.Cells;

        for (int index = 0; index < copied.Count; index++)
        {
            string text = string.Empty;

            if (index < cells.Length)
            {
                text = cells[index];
            }

            copied[index].SetText(text);
        }

        return copy;
    }
}

public sealed partial class TableRow
{
    /// <summary>
    /// Adds a copy of the row after it, with its formatting and content; a header row's copy is
    /// not a header. For a template row to fill in again and again.
    /// </summary>
    /// <returns>The copy.</returns>
    /// <exception cref="InvalidOperationException">The row was read with <see cref="DocxReader"/>.</exception>
    public TableRow InsertCopyAfter()
    {
        Context.ThrowIfReadOnly();
        XElement copy = new(Element);
        copy.Element(Context.Names.TableRowProperties)?.Element(Context.Names.W + "tblHeader")?.Remove();

        // The copy is the row as it reads now; the original's tracked changes, and their ids, are not copied.
        Revisions.Accept(Context.Names, copy);
        Element.AddAfterSelf(copy);
        Context.Owner?.Tracker?.InsertBlock(Context.Names, copy);
        Context.Changed();

        return new TableRow(Context, copy, Table);
    }

    /// <summary>Takes the row out of its table; while changes are tracked, marks it deleted instead.</summary>
    /// <exception cref="InvalidOperationException">The row was read with <see cref="DocxReader"/>.</exception>
    public void Remove()
    {
        Context.ThrowIfReadOnly();

        if (Context.Owner?.Tracker is Tracking tracking)
        {
            tracking.DeleteBlock(Context.Names, Element);
        }
        else
        {
            Element.Remove();
        }

        Context.Changed();
    }
}

public sealed partial class TableCell
{
    /// <summary>Replaces the cell's content with one paragraph of text, in the formatting of its first paragraph and run.</summary>
    /// <exception cref="InvalidOperationException">The cell was read with <see cref="DocxReader"/>.</exception>
    public void SetText(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        Context.ThrowIfReadOnly();
        WordNames names = Context.Names;
        XElement? first = Element.Elements(names.Paragraph).FirstOrDefault();
        XElement paragraph = new(names.Paragraph);

        if (first?.Element(names.ParagraphProperties) is XElement properties)
        {
            XElement copy = new(properties);
            Revisions.Accept(names, copy);
            paragraph.Add(copy);
        }

        XElement? runProperties = first?.Descendants(names.Run).Select(run => run.Element(names.RunProperties)).FirstOrDefault(found => found is not null);
        XElement run = new(names.Run);

        if (runProperties is not null)
        {
            run.Add(new XElement(runProperties));
        }

        ContentXml.AddText(names, run, text);
        paragraph.Add(run);
        Tracking? tracking = Context.Owner?.Tracker;

        foreach (XElement block in Element.Elements().Where(element => element.Name != names.TableCellProperties).ToList())
        {
            if (tracking is null)
            {
                block.Remove();
            }
            else
            {
                tracking.DeleteBlock(names, block);
            }
        }

        Element.Add(paragraph);
        tracking?.InsertBlock(names, paragraph);
        Context.Changed();
    }
}
