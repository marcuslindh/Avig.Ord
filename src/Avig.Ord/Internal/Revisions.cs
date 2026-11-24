using System.Xml.Linq;

namespace Avig.Ord.Internal;

/// <summary>Accepts or rejects every tracked change under an element.</summary>
internal static class Revisions
{
    private static readonly string[] PropertyChanges = ["rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange", "sectPrChange", "tblGridChange", "numberingChange"];

    public static void Accept(WordNames names, XElement root)
    {
        // Deleted content goes, with whole rows marked deleted.
        Remove(root.Descendants().Where(element => element.Name == names.Deleted && IsContentRevision(names, element) == true));
        Remove(root.Descendants(names.MoveFrom));
        Remove(root.Descendants(names.TableRow).Where(row => row.Element(names.TableRowProperties)?.Element(names.Deleted) is not null));
        Remove(root.Descendants().Where(element => element.Name.LocalName is "moveFromRangeStart" or "moveFromRangeEnd" or "moveToRangeStart" or "moveToRangeEnd" or "cellDel"));

        // Inserted content stays, as ordinary content.
        Unwrap(root.Descendants().Where(element => (element.Name == names.Inserted && IsContentRevision(names, element) == true) || element.Name == names.MoveTo));

        // A deleted paragraph mark joins the paragraph with the next.
        JoinParagraphs(names, root, names.Deleted);

        // The marks left in properties (an inserted paragraph mark, an inserted row) and the old properties of formatting changes.
        Remove(root.Descendants().Where(element => (element.Name == names.Inserted || element.Name == names.Deleted || element.Name.LocalName == "cellIns" || element.Name.LocalName == "cellMerge") && IsContentRevision(names, element) == false));
        Remove(root.Descendants().Where(element => element.Name.Namespace == names.W && PropertyChanges.Contains(element.Name.LocalName) == true));
    }

    public static void Reject(WordNames names, XElement root)
    {
        // Inserted content goes, with whole rows marked inserted.
        Remove(root.Descendants().Where(element => element.Name == names.Inserted && IsContentRevision(names, element) == true));
        Remove(root.Descendants(names.MoveTo));
        Remove(root.Descendants(names.TableRow).Where(row => row.Element(names.TableRowProperties)?.Element(names.Inserted) is not null));
        Remove(root.Descendants().Where(element => element.Name.LocalName is "moveFromRangeStart" or "moveFromRangeEnd" or "moveToRangeStart" or "moveToRangeEnd" or "cellIns"));

        // Deleted content comes back as ordinary text.
        foreach (XElement text in root.Descendants(names.DeletedText).ToList())
        {
            text.Name = names.Text;
        }

        foreach (XElement instruction in root.Descendants(names.DeletedInstructionText).ToList())
        {
            instruction.Name = names.InstructionText;
        }

        Unwrap(root.Descendants().Where(element => (element.Name == names.Deleted && IsContentRevision(names, element) == true) || element.Name == names.MoveFrom));

        // An inserted paragraph mark was not there before: the paragraph joins the next.
        JoinParagraphs(names, root, names.Inserted);

        // Formatting changes go back to the properties they had.
        foreach (XElement change in root.Descendants().Where(element => element.Name.Namespace == names.W && PropertyChanges.Contains(element.Name.LocalName) == true).ToList())
        {
            XElement? properties = change.Parent;
            XElement? old = change.Elements().FirstOrDefault();

            if (properties is null)
            {
                continue;
            }

            change.Remove();

            if (old is null)
            {
                continue;
            }

            // The old properties replace the current ones, but what is not a property (a
            // section break, the paragraph mark's run properties) stays.
            foreach (XElement current in properties.Elements().Where(element => element.Name.LocalName is not ("sectPr" or "rPr") || properties.Name.LocalName == "rPr").ToList())
            {
                current.Remove();
            }

            properties.AddFirst(old.Elements());
        }

        Remove(root.Descendants().Where(element => (element.Name == names.Inserted || element.Name == names.Deleted || element.Name.LocalName == "cellDel" || element.Name.LocalName == "cellMerge") && IsContentRevision(names, element) == false));
    }

    /// <summary>
    /// Joins every paragraph whose mark has the revision <paramref name="mark"/> with the
    /// paragraph after it: its content moves to the start of the next, which keeps its own
    /// properties. A paragraph with no paragraph after it in its container keeps its mark.
    /// </summary>
    private static void JoinParagraphs(WordNames names, XElement root, XName mark)
    {
        foreach (XElement paragraph in root.Descendants(names.Paragraph).Reverse().ToList())
        {
            XElement? revision = paragraph.Element(names.ParagraphProperties)?.Element(names.RunProperties)?.Element(mark);

            if (revision is null || paragraph.Parent is null)
            {
                continue;
            }

            if (paragraph.ElementsAfterSelf().FirstOrDefault() is not XElement next || next.Name != names.Paragraph)
            {
                // With nothing left in it and something after it (a table), the paragraph simply goes.
                bool empty = paragraph.Descendants(names.Run).Any() == false;
                bool endsSection = paragraph.Element(names.ParagraphProperties)?.Element(names.SectionProperties) is not null;

                if (empty == true && endsSection == false && paragraph.ElementsAfterSelf().Any() == true)
                {
                    paragraph.Remove();
                }

                continue;
            }

            List<XNode> content = [.. paragraph.Nodes().Where(node => node is not XElement element || element.Name != names.ParagraphProperties)];

            foreach (XNode node in content)
            {
                node.Remove();
            }

            XElement? nextProperties = next.Element(names.ParagraphProperties);

            if (nextProperties is null)
            {
                next.AddFirst(content);
            }
            else
            {
                nextProperties.AddAfterSelf(content);
            }

            // A section break on the joined paragraph's mark goes with it.
            XElement? section = paragraph.Element(names.ParagraphProperties)?.Element(names.SectionProperties);

            if (section is not null && next.Element(names.ParagraphProperties)?.Element(names.SectionProperties) is null)
            {
                section.Remove();

                if (nextProperties is null)
                {
                    nextProperties = new XElement(names.ParagraphProperties);
                    next.AddFirst(nextProperties);
                }

                nextProperties.Add(section);
            }

            paragraph.Remove();
        }
    }

    /// <summary>A w:ins or w:del that holds content, as opposed to the mark in a paragraph mark's or a row's properties.</summary>
    private static bool IsContentRevision(WordNames names, XElement element)
    {
        XName? parent = element.Parent?.Name;

        return parent != names.RunProperties && parent != names.TableRowProperties && parent != names.TableCellProperties && parent != names.W + "numPr";
    }

    private static void Remove(IEnumerable<XElement> elements)
    {
        foreach (XElement element in elements.ToList())
        {
            if (element.Parent is not null)
            {
                element.Remove();
            }
        }
    }

    private static void Unwrap(IEnumerable<XElement> elements)
    {
        // Innermost first, so that a revision inside another is unwrapped before its parent moves.
        foreach (XElement element in elements.Reverse().ToList())
        {
            if (element.Parent is not null)
            {
                element.ReplaceWith(element.Nodes());
            }
        }
    }
}
