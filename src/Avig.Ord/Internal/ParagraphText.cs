using System.Text;
using System.Xml.Linq;

namespace Avig.Ord.Internal;

/// <summary>
/// A paragraph's visible text (as the Accepted view has it) with where each character comes
/// from: which w:t (and where in it), or which w:tab, w:br or w:cr. Search and replace work on
/// the text and then split and change the runs it maps to, so that a match can span runs that
/// Word split (spell-check marks, formatting, revision ids) and keep the formatting it had.
/// </summary>
internal sealed class ParagraphText
{
    private readonly WordNames names;
    private readonly List<Segment> segments = [];

    private ParagraphText(WordNames names, XElement paragraph)
    {
        this.names = names;
        Paragraph = paragraph;
        StringBuilder text = new();
        Collect(paragraph, text, new Stack<bool>(), 0);
        Text = text.ToString();
    }

    public XElement Paragraph { get; }

    public string Text { get; }

    public static ParagraphText Read(WordNames names, XElement paragraph)
    {
        return new ParagraphText(names, paragraph);
    }

    /// <summary>
    /// Splits runs so that a run starts at <paramref name="start"/> and one ends at
    /// <paramref name="end"/>, and gives back the runs between, in order. The text read after
    /// this is the same; only the runs are cut.
    /// </summary>
    public static List<XElement> Isolate(WordNames names, XElement paragraph, int start, int end)
    {
        SplitAt(names, paragraph, end);
        SplitAt(names, paragraph, start);
        ParagraphText text = Read(names, paragraph);
        List<XElement> runs = [];

        foreach (Segment segment in text.segments)
        {
            if (segment.Start >= start && segment.Start + segment.Length <= end && segment.Length > 0 && runs.Contains(segment.Run) == false)
            {
                runs.Add(segment.Run);
            }
        }

        return runs;
    }

    /// <summary>The run the character at <paramref name="position"/> is in, or the last run before it.</summary>
    public XElement? RunAt(int position)
    {
        XElement? run = null;

        foreach (Segment segment in segments)
        {
            if (segment.Start <= position)
            {
                run = segment.Run;
            }
        }

        return run;
    }

    private static void SplitAt(WordNames names, XElement paragraph, int position)
    {
        ParagraphText text = Read(names, paragraph);

        foreach (Segment segment in text.segments)
        {
            if (position <= segment.Start || position >= segment.Start + segment.Length)
            {
                // Not inside this segment; but a boundary between two segments of one run is a split too.
                if (position == segment.Start && segment.Start > 0)
                {
                    Segment? previous = text.segments.LastOrDefault(other => other.Start + other.Length == position && other.Length > 0);

                    if (previous is not null && previous.Run == segment.Run)
                    {
                        SplitRunBefore(names, segment.Run, segment.Node, 0);

                        return;
                    }
                }

                continue;
            }

            if (segment.Node.Name == names.Text)
            {
                SplitRunBefore(names, segment.Run, segment.Node, position - segment.Start);
            }

            return;
        }
    }

    /// <summary>
    /// Cuts a run in two before <paramref name="node"/>, or inside it at <paramref name="offset"/>
    /// when it is a w:t: everything from there on moves to a new run with a copy of the run's
    /// properties, right after it.
    /// </summary>
    private static void SplitRunBefore(WordNames names, XElement run, XElement node, int offset)
    {
        XElement second = new(names.Run);

        foreach (XAttribute attribute in run.Attributes())
        {
            second.SetAttributeValue(attribute.Name, attribute.Value);
        }

        XElement? properties = run.Element(names.RunProperties);

        if (properties is not null)
        {
            second.Add(new XElement(properties));
        }

        List<XNode> moving = [];

        if (offset > 0 && node.Name == names.Text)
        {
            string value = node.Value;
            XElement tail = new(names.Text, value[offset..]);
            tail.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            node.Value = value[..offset];
            node.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            second.Add(tail);
            moving.AddRange(node.NodesAfterSelf());
        }
        else
        {
            moving.Add(node);
            moving.AddRange(node.NodesAfterSelf());
        }

        foreach (XNode moved in moving)
        {
            moved.Remove();
            second.Add(moved);
        }

        run.AddAfterSelf(second);
    }

    private void Collect(XElement container, StringBuilder text, Stack<bool> fields, int depth)
    {
        foreach (XElement child in container.Elements())
        {
            XName name = child.Name;

            if (name == names.Run)
            {
                CollectRun(child, text, fields);
            }
            else if (name == names.Deleted || name == names.MoveFrom || name == names.ParagraphProperties)
            {
                continue;
            }
            else if (depth < 16 && (name == names.Hyperlink || name == names.Inserted || name == names.MoveTo || name == names.SimpleField || name == names.W + "smartTag" || name == names.W + "customXml" || name == names.W + "bdo" || name == names.W + "dir"))
            {
                Collect(child, text, fields, depth + 1);
            }
            else if (depth < 16 && name == names.ContentControl && child.Element(names.ContentControlContent) is XElement content)
            {
                Collect(content, text, fields, depth + 1);
            }
        }
    }

    private void CollectRun(XElement run, StringBuilder text, Stack<bool> fields)
    {
        foreach (XElement child in run.Elements())
        {
            if (child.Name.Namespace != names.W)
            {
                continue;
            }

            switch (child.Name.LocalName)
            {
                case "fldChar":
                    string? type = child.Attribute(names.W + "fldCharType")?.Value;

                    if (type == "begin")
                    {
                        fields.Push(true);
                    }
                    else if (type == "separate" && fields.Count > 0)
                    {
                        fields.Pop();
                        fields.Push(false);
                    }
                    else if (type == "end" && fields.Count > 0)
                    {
                        fields.Pop();
                    }

                    break;

                case "t":
                    Add(run, child, child.Value, text, fields);

                    break;

                case "tab":
                case "ptab":
                    Add(run, child, "\t", text, fields);

                    break;

                case "br":
                case "cr":
                    if (child.Attribute(names.Type)?.Value is not ("page" or "column"))
                    {
                        Add(run, child, "\n", text, fields);
                    }

                    break;

                default:
                    break;
            }
        }
    }

    private void Add(XElement run, XElement node, string value, StringBuilder text, Stack<bool> fields)
    {
        // Text inside a field's instruction is not shown; its result is.
        if (fields.Contains(true) == true || value.Length == 0)
        {
            return;
        }

        segments.Add(new Segment(run, node, text.Length, value.Length));
        text.Append(value);
    }

    private sealed record Segment(XElement Run, XElement Node, int Start, int Length);
}
