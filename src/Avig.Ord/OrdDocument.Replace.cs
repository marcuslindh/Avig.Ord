using System.Text.RegularExpressions;
using System.Xml.Linq;
using Avig.Ord.Internal;

namespace Avig.Ord;

public sealed partial class OrdDocument
{
    private Tracking? tracking;

    /// <summary>True while changes are tracked (see <see cref="TrackChanges"/>).</summary>
    public bool IsTrackingChanges => tracking is not null;

    internal Tracking? Tracker => tracking;

    /// <summary>
    /// From now on, changes are saved as tracked changes with this author: what is replaced or
    /// removed is marked deleted and what is added marked inserted, so that the changes can be
    /// read, accepted or rejected in Word. It applies to <see cref="Replace(string, Replacement, SearchOptions?)"/>,
    /// template filling and the editing methods on paragraphs, tables and rows.
    /// </summary>
    /// <param name="author">The name Word shows for the changes.</param>
    /// <param name="date">When the changes were made; now when null.</param>
    public void TrackChanges(string author, DateTimeOffset? date = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(author);
        tracking = new Tracking(author, date ?? DateTimeOffset.UtcNow, Tracking.HighestId(ContentRoots()));
    }

    /// <summary>Changes are made straight from now on, not tracked.</summary>
    public void StopTrackingChanges()
    {
        tracking = null;
    }

    /// <summary>Every place a text is found, paragraph by paragraph. A match never goes past its paragraph.</summary>
    public IReadOnlyList<TextMatch> Find(string text, SearchOptions? options = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(text);
        options ??= SearchOptions.Default;
        List<TextMatch> found = [];

        foreach ((DocumentContext context, XElement paragraph) in SearchParagraphs(options.Scope))
        {
            string value = ParagraphText.Read(context.Names, paragraph).Text;

            foreach ((int index, int length) in Occurrences(value, text, options))
            {
                found.Add(new TextMatch(new Paragraph(context, paragraph), index, value.Substring(index, length), null));
            }
        }

        return found;
    }

    /// <summary>Every match of a regular expression, paragraph by paragraph.</summary>
    public IReadOnlyList<TextMatch> Find(Regex pattern, SearchOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(pattern);
        options ??= SearchOptions.Default;
        List<TextMatch> found = [];

        foreach ((DocumentContext context, XElement paragraph) in SearchParagraphs(options.Scope))
        {
            string value = ParagraphText.Read(context.Names, paragraph).Text;

            foreach (Match match in pattern.Matches(value).Where(match => match.Length > 0 && IsWholeWord(value, match.Index, match.Length, options) == true))
            {
                found.Add(new TextMatch(new Paragraph(context, paragraph), match.Index, match.Value, match));
            }
        }

        return found;
    }

    /// <summary>
    /// Replaces every occurrence of a text, also where Word has split it over several runs,
    /// keeping the formatting of the run each occurrence starts in.
    /// </summary>
    /// <returns>How many were replaced.</returns>
    /// <example>
    /// <code>
    /// int count = document.Replace("2025", "2026");
    /// document.Replace("[DATE]", Replacement.Text(today, s =&gt; s.Bold()));
    /// </code>
    /// </example>
    public int Replace(string find, Replacement replacement, SearchOptions? options = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(find);
        ArgumentNullException.ThrowIfNull(replacement);
        options ??= SearchOptions.Default;

        return ReplaceAll(options, value => Occurrences(value, find, options).Select(occurrence => (occurrence.Index, occurrence.Length, replacement)));
    }

    /// <summary>Replaces every match of a regular expression; <paramref name="replacement"/> may refer to groups as $1 or ${name}.</summary>
    /// <returns>How many were replaced.</returns>
    public int Replace(Regex pattern, string replacement, SearchOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(pattern);
        ArgumentNullException.ThrowIfNull(replacement);

        return Replace(pattern, match => Replacement.Text(match.Result(replacement)), options);
    }

    /// <summary>Replaces every match of a regular expression with what <paramref name="replacement"/> makes of it: text, a picture, paragraphs or a table.</summary>
    /// <returns>How many were replaced.</returns>
    public int Replace(Regex pattern, Func<Match, Replacement> replacement, SearchOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(pattern);
        ArgumentNullException.ThrowIfNull(replacement);
        options ??= SearchOptions.Default;

        return ReplaceAll(options, value => pattern.Matches(value)
            .Where(match => match.Length > 0 && IsWholeWord(value, match.Index, match.Length, options) == true)
            .Select(match => (match.Index, match.Length, replacement(match))));
    }

    /// <summary>Replaces matches the caller finds in each paragraph's text: (index, length, replacement), in order and not overlapping.</summary>
    internal int ReplaceAll(SearchOptions options, Func<string, IEnumerable<(int Index, int Length, Replacement Replacement)>> matches)
    {
        int count = 0;

        foreach ((DocumentContext context, XElement paragraph) in SearchParagraphs(options.Scope))
        {
            if (paragraph.Parent is null)
            {
                continue;
            }

            string value = ParagraphText.Read(context.Names, paragraph).Text;
            List<(int Index, int Length, Replacement Replacement)> found = [.. matches(value)];

            if (found.Count == 0)
            {
                continue;
            }

            // From the end, so that the positions of the matches before stay right.
            for (int index = found.Count - 1; index >= 0; index--)
            {
                (int start, int length, Replacement replacement) = found[index];

                if (ReplaceIn(context, paragraph, start, start + length, replacement) == true)
                {
                    count++;
                }
            }

            context.Changed();
        }

        return count;
    }

    private static IEnumerable<(int Index, int Length)> Occurrences(string value, string find, SearchOptions options)
    {
        StringComparison comparison = StringComparison.Ordinal;

        if (options.MatchCase == false)
        {
            comparison = StringComparison.OrdinalIgnoreCase;
        }

        int index = value.IndexOf(find, comparison);

        while (index >= 0)
        {
            if (IsWholeWord(value, index, find.Length, options) == true)
            {
                yield return (index, find.Length);
                index = value.IndexOf(find, index + find.Length, comparison);
            }
            else
            {
                index = value.IndexOf(find, index + 1, comparison);
            }
        }
    }

    private static bool IsWholeWord(string value, int index, int length, SearchOptions options)
    {
        if (options.WholeWord == false)
        {
            return true;
        }

        bool before = index > 0 && char.IsLetterOrDigit(value[index - 1]) == true;
        bool after = index + length < value.Length && char.IsLetterOrDigit(value[index + length]) == true;

        return before == false && after == false;
    }

    /// <summary>Removes hyperlinks, insertions and content control content that changes left without a run.</summary>
    private static void RemoveEmptyContainers(WordNames names, XElement paragraph)
    {
        foreach (XElement container in paragraph.Descendants().Where(element => element.Name == names.Hyperlink || element.Name == names.Inserted || element.Name == names.W + "smartTag").Reverse().ToList())
        {
            if (container.Descendants(names.Run).Any() == false)
            {
                container.Remove();
            }
        }
    }

    private static bool IsEmpty(WordNames names, XElement paragraph)
    {
        return paragraph.Descendants(names.Run).Any() == false;
    }

    /// <summary>The paragraphs to search, part by part, each with its context; a snapshot, since replacing adds and removes paragraphs.</summary>
    private List<(DocumentContext Context, XElement Paragraph)> SearchParagraphs(SearchScope scope)
    {
        List<(DocumentContext Context, XElement Paragraph)> paragraphs = [];

        void Add(string part, XElement root)
        {
            DocumentContext context = ContextFor(part);
            paragraphs.AddRange(root.Descendants(context.Names.Paragraph).Select(paragraph => (context, paragraph)));
        }

        if (scope.HasFlag(SearchScope.Body) == true)
        {
            Add(MainPartName, Body);
        }

        if (scope.HasFlag(SearchScope.HeadersAndFooters) == true)
        {
            foreach (HeaderFooter header in Headers.Concat(Footers))
            {
                Add(header.PartName, header.Root);
            }
        }

        List<string> kinds = [];

        if (scope.HasFlag(SearchScope.Notes) == true)
        {
            kinds.Add("footnotes");
            kinds.Add("endnotes");
        }

        if (scope.HasFlag(SearchScope.Comments) == true)
        {
            kinds.Add("comments");
        }

        foreach (string kind in kinds)
        {
            string? part = RelatedPart(MainPartName, kind);

            if (part is not null && Xml(part).Root is XElement root)
            {
                Add(part, root);
            }
        }

        return paragraphs;
    }

    private bool ReplaceIn(DocumentContext context, XElement paragraph, int start, int end, Replacement replacement)
    {
        WordNames names = context.Names;
        List<XElement> runs = ParagraphText.Isolate(names, paragraph, start, end);

        if (runs.Count == 0)
        {
            return false;
        }

        if (replacement.Kind == ReplacementKind.Blocks)
        {
            ReplaceWithBlocks(context, paragraph, runs, replacement);

            return true;
        }

        XElement? properties = runs[0].Element(names.RunProperties);
        XElement? added = MakeInline(context, properties, replacement);

        if (tracking is not null)
        {
            tracking.Delete(names, runs);
            XElement anchor = runs[^1];

            if (anchor.Parent?.Name == names.Deleted)
            {
                anchor = anchor.Parent;
            }

            if (added is not null)
            {
                anchor.AddAfterSelf(added);
                tracking.Insert(names, [added]);
            }
        }
        else
        {
            if (added is not null)
            {
                runs[0].AddBeforeSelf(added);
            }

            foreach (XElement run in runs)
            {
                run.Remove();
            }

            RemoveEmptyContainers(names, paragraph);
        }

        return true;
    }

    /// <summary>The run a text or picture replacement becomes; null for an empty text.</summary>
    private XElement? MakeInline(DocumentContext context, XElement? properties, Replacement replacement)
    {
        WordNames names = context.Names;
        IContentSink sink = SinkFor(context);
        XElement? made = null;

        if (replacement.Kind == ReplacementKind.Picture)
        {
            ParagraphBuilder builder = new();

            if (replacement.PictureWidth is Length width && replacement.PictureHeight is Length height)
            {
                builder.Picture(replacement.PictureData!, width, height, replacement.Description);
            }
            else
            {
                builder.Picture(replacement.PictureData!, replacement.Description);
            }

            made = builder.Build(sink).Element(names.Run);
            made?.Remove();
        }
        else if (replacement.TextValue.Length > 0)
        {
            made = new XElement(names.Run);
            XElement? merged = properties;

            if (replacement.Style is not null)
            {
                RunStyle style = new();
                replacement.Style(style);
                merged = RunProperties.Merge(names, properties, style.ToProperties(sink));
            }

            if (merged is not null)
            {
                made.Add(new XElement(merged));
            }

            ContentXml.AddText(names, made, replacement.TextValue);
        }

        return made;
    }

    /// <summary>
    /// Splits the paragraph where the match is: the text before stays, the new blocks come
    /// next, then a paragraph with the text after, in the paragraph's own formatting. A part
    /// that is left empty goes, unless it ends a section.
    /// </summary>
    private void ReplaceWithBlocks(DocumentContext context, XElement paragraph, List<XElement> runs, Replacement replacement)
    {
        WordNames names = context.Names;
        IContentSink sink = SinkFor(context);
        XElement? properties = paragraph.Element(names.ParagraphProperties);
        XElement after = new(names.Paragraph);

        if (properties is not null)
        {
            XElement copy = new(properties);
            copy.Element(names.RunProperties)?.Element(names.Inserted)?.Remove();
            after.Add(copy);

            // The section break belongs to the last of the paragraphs.
            properties.Element(names.SectionProperties)?.Remove();
        }

        XElement last = TopLevel(paragraph, runs[^1]);
        List<XNode> moving = [.. last.NodesAfterSelf()];

        if (tracking is not null)
        {
            tracking.Delete(names, runs);
        }
        else
        {
            foreach (XElement run in runs)
            {
                run.Remove();
            }
        }

        if (last.Parent == paragraph && last.Name != names.Run && last.Descendants(names.Run).Any() == true && runs.Contains(last) == false)
        {
            moving.Insert(0, last);
        }

        foreach (XNode node in moving)
        {
            node.Remove();
            after.Add(node);
        }

        List<XElement> blocks = [];

        foreach (Action<ParagraphBuilder> build in replacement.ParagraphBuilders)
        {
            ParagraphBuilder builder = new();
            build(builder);
            blocks.Add(builder.Build(sink));
        }

        if (replacement.TableBuilder is not null)
        {
            TableBuilder builder = new();
            replacement.TableBuilder(builder);
            blocks.Add(builder.Build(sink, sink.TextWidth));
        }

        XElement anchor = paragraph;

        foreach (XElement block in blocks)
        {
            anchor.AddAfterSelf(block);
            anchor = block;
            tracking?.InsertBlock(names, block);
        }

        anchor.AddAfterSelf(after);
        RemoveEmptyContainers(names, paragraph);
        RemoveEmptyContainers(names, after);
        bool keepAfter = after.Element(names.ParagraphProperties)?.Element(names.SectionProperties) is not null || IsEmpty(names, after) == false;

        if (tracking is not null)
        {
            tracking.MarkParagraph(names, paragraph, true);
        }
        else if (IsEmpty(names, paragraph) == true)
        {
            paragraph.Remove();
        }

        if (keepAfter == false)
        {
            after.Remove();
        }

        // A table cell must end with a paragraph.
        XElement? cell = anchor.Parent;

        if (cell is not null && cell.Name == names.TableCell && cell.Elements().LastOrDefault()?.Name != names.Paragraph)
        {
            cell.Add(new XElement(names.Paragraph));
        }
    }

    /// <summary>The child of the paragraph that holds <paramref name="run"/>: the run itself, or the hyperlink, field or content control it is in.</summary>
    private static XElement TopLevel(XElement paragraph, XElement run)
    {
        XElement current = run;

        while (current.Parent is not null && current.Parent != paragraph)
        {
            current = current.Parent;
        }

        return current;
    }
}
