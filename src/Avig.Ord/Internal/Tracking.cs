using System.Globalization;
using System.Xml.Linq;

namespace Avig.Ord.Internal;

/// <summary>Marks content as inserted or deleted, as Word does when it tracks changes.</summary>
internal sealed class Tracking
{
    private readonly string author;
    private readonly string date;
    private int nextId;

    public Tracking(string author, DateTimeOffset date, int firstId)
    {
        this.author = author;
        this.date = date.UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
        nextId = firstId;
    }

    public string Author => author;

    /// <summary>The highest w:id a revision, comment or bookmark has under the roots, so that new revisions get ids of their own.</summary>
    public static int HighestId(IEnumerable<XElement> roots)
    {
        int highest = 0;

        foreach (XElement root in roots)
        {
            foreach (XElement element in root.Descendants())
            {
                XAttribute? id = element.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == "id" && attribute.Name.NamespaceName.Contains("wordprocessingml", StringComparison.Ordinal) == true);

                if (id is not null && int.TryParse(id.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) == true && value > highest)
                {
                    highest = value;
                }
            }
        }

        return highest;
    }

    /// <summary>A w:ins or w:del with an id, the author and the date.</summary>
    public XElement Mark(WordNames names, bool inserted)
    {
        XName name = names.Deleted;

        if (inserted == true)
        {
            name = names.Inserted;
        }

        nextId++;

        return new XElement(name, new XAttribute(names.WId, nextId.ToString(CultureInfo.InvariantCulture)), new XAttribute(names.W + "author", author), new XAttribute(names.W + "date", date));
    }

    /// <summary>Wraps runs in a w:ins, one wrapper around the runs that sit next to each other.</summary>
    public void Insert(WordNames names, IReadOnlyList<XElement> runs)
    {
        foreach (XElement run in runs)
        {
            if (run.Parent is null || run.Parent.Name == names.Inserted)
            {
                continue;
            }

            XElement wrapper = Mark(names, true);
            run.AddBeforeSelf(wrapper);
            run.Remove();
            wrapper.Add(run);
        }
    }

    /// <summary>Marks runs deleted: their text becomes w:delText, and each is wrapped in a w:del. Runs that were inserted in a tracked change are simply removed, as Word does.</summary>
    public void Delete(WordNames names, IReadOnlyList<XElement> runs)
    {
        foreach (XElement run in runs)
        {
            if (run.Parent is null)
            {
                continue;
            }

            if (run.Parent.Name == names.Inserted)
            {
                XElement insertion = run.Parent;
                run.Remove();

                if (insertion.HasElements == false)
                {
                    insertion.Remove();
                }

                continue;
            }

            if (run.Parent.Name == names.Deleted)
            {
                continue;
            }

            foreach (XElement text in run.Elements(names.Text).ToList())
            {
                text.Name = names.DeletedText;
            }

            foreach (XElement instruction in run.Elements(names.InstructionText).ToList())
            {
                instruction.Name = names.DeletedInstructionText;
            }

            XElement wrapper = Mark(names, false);
            run.AddBeforeSelf(wrapper);
            run.Remove();
            wrapper.Add(run);
        }
    }

    /// <summary>Marks a new block inserted: its runs, its paragraph marks and, in a table, its rows.</summary>
    public void InsertBlock(WordNames names, XElement block)
    {
        foreach (XElement paragraph in ParagraphsOf(names, block))
        {
            Insert(names, [.. paragraph.Descendants(names.Run).Where(run => run.Ancestors(names.Paragraph).FirstOrDefault() == paragraph)]);
            MarkParagraph(names, paragraph, true);
        }

        foreach (XElement row in block.DescendantsAndSelf(names.TableRow))
        {
            MarkRow(names, row, true);
        }
    }

    /// <summary>
    /// Marks a block deleted instead of removing it. What was itself inserted while changes
    /// were tracked (a paragraph or row whose mark is an insertion) is removed outright, as Word
    /// does, since there is nothing to go back to.
    /// </summary>
    public void DeleteBlock(WordNames names, XElement block)
    {
        foreach (XElement row in block.DescendantsAndSelf(names.TableRow).ToList())
        {
            if (row.Element(names.TableRowProperties)?.Element(names.Inserted) is not null)
            {
                row.Remove();
            }
        }

        foreach (XElement paragraph in ParagraphsOf(names, block))
        {
            Delete(names, [.. paragraph.Descendants(names.Run).Where(run => run.Ancestors(names.Paragraph).FirstOrDefault() == paragraph)]);
            bool inserted = paragraph.Element(names.ParagraphProperties)?.Element(names.RunProperties)?.Element(names.Inserted) is not null;

            if (inserted == true && paragraph.Descendants(names.Run).Any() == false)
            {
                paragraph.Remove();

                continue;
            }

            MarkParagraph(names, paragraph, false);
        }

        foreach (XElement row in block.DescendantsAndSelf(names.TableRow))
        {
            MarkRow(names, row, false);
        }
    }

    /// <summary>Marks a paragraph's mark inserted or deleted (w:pPr/w:rPr/w:ins), which joins or splits paragraphs when the change is accepted.</summary>
    public void MarkParagraph(WordNames names, XElement paragraph, bool inserted)
    {
        XElement? properties = paragraph.Element(names.ParagraphProperties);

        if (properties is null)
        {
            properties = new XElement(names.ParagraphProperties);
            paragraph.AddFirst(properties);
        }

        XElement? markProperties = properties.Element(names.RunProperties);

        if (markProperties is null)
        {
            markProperties = new XElement(names.RunProperties);
            XElement? after = properties.Elements().FirstOrDefault(element => element.Name == names.SectionProperties || element.Name.LocalName == "pPrChange");

            if (after is null)
            {
                properties.Add(markProperties);
            }
            else
            {
                after.AddBeforeSelf(markProperties);
            }
        }

        if (markProperties.Element(names.Inserted) is null && markProperties.Element(names.Deleted) is null)
        {
            markProperties.AddFirst(Mark(names, inserted));
        }
    }

    public void MarkRow(WordNames names, XElement row, bool inserted)
    {
        XElement? properties = row.Element(names.TableRowProperties);

        if (properties is null)
        {
            properties = new XElement(names.TableRowProperties);
            XElement? exceptions = row.Element(names.W + "tblPrEx");

            if (exceptions is null)
            {
                row.AddFirst(properties);
            }
            else
            {
                exceptions.AddAfterSelf(properties);
            }
        }

        if (properties.Element(names.Inserted) is null && properties.Element(names.Deleted) is null)
        {
            XElement? change = properties.Element(names.W + "trPrChange");
            XElement mark = Mark(names, inserted);

            if (change is null)
            {
                properties.Add(mark);
            }
            else
            {
                change.AddBeforeSelf(mark);
            }
        }
    }

    private static IEnumerable<XElement> ParagraphsOf(WordNames names, XElement block)
    {
        return [.. block.DescendantsAndSelf(names.Paragraph)];
    }
}
