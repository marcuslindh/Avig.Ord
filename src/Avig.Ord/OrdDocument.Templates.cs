using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Avig.Ord.Internal;

namespace Avig.Ord;

public sealed partial class OrdDocument
{
    /// <summary>How deep sections may be nested in a template, so that a template cannot make the filling recurse without end.</summary>
    private const int MaxSectionDepth = 32;

    /// <summary>
    /// Fills the document as a template, in the body, headers, footers and notes:
    /// <list type="bullet">
    /// <item><c>{{Name}}</c> becomes the value, <c>{{Total:N2}}</c> and <c>{{Date:yyyy-MM-dd}}</c> formatted in <see cref="TemplateOptions.Culture"/>, <c>{{Customer.Name}}</c> a nested value.</item>
    /// <item><c>{{#Lines}}</c> … <c>{{/Lines}}</c> repeats what is between for each item of a list (table rows when the tags are in rows of a table, paragraphs when they are in paragraphs of their own), and keeps it once or not at all for true or false; <c>{{^Name}}</c> … <c>{{/Name}}</c> keeps it when the value is false, empty or missing. Inside, <c>{{.}}</c> is the item itself.</item>
    /// <item>A <see cref="Replacement"/> as the value puts in a picture, paragraphs or a table.</item>
    /// <item>MERGEFIELD fields and content controls (by tag, or title) get the value of their name.</item>
    /// </list>
    /// A placeholder without a value becomes empty and is listed in <see cref="TemplateResult.Missing"/>.
    /// </summary>
    /// <returns>How many were filled, and those that had no value.</returns>
    public TemplateResult FillTemplate(ITemplateData data, TemplateOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(data);
        options ??= TemplateOptions.Default;
        TemplateFill fill = new(this, options);
        TemplateScope scope = new(data);

        foreach ((DocumentContext context, XElement container) in TemplateContainers())
        {
            fill.Blocks(context, container, scope, 0);
            context.Changed();
        }

        return new TemplateResult(fill.Filled, fill.Missing);
    }

    /// <summary>Fills the document as a template with values by name (see <see cref="FillTemplate(ITemplateData, TemplateOptions?)"/>).</summary>
    /// <returns>How many were filled, and those that had no value.</returns>
    public TemplateResult FillTemplate(IEnumerable<KeyValuePair<string, object?>> values, TemplateOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(values);

        return FillTemplate(TemplateValues.From(values), options);
    }

    private List<(DocumentContext Context, XElement Container)> TemplateContainers()
    {
        List<(DocumentContext Context, XElement Container)> containers = [(MainContext, Body)];

        foreach (HeaderFooter header in Headers.Concat(Footers))
        {
            containers.Add((header.Context, header.Root));
        }

        foreach (string kind in (string[])["footnotes", "endnotes"])
        {
            string? part = RelatedPart(MainPartName, kind);

            if (part is not null && Xml(part).Root is XElement root)
            {
                DocumentContext context = ContextFor(part);

                foreach (XElement note in root.Elements())
                {
                    containers.Add((context, note));
                }
            }
        }

        return containers;
    }

    /// <summary>The work of one <see cref="FillTemplate(ITemplateData, TemplateOptions?)"/>.</summary>
    private sealed partial class TemplateFill
    {
        private readonly OrdDocument document;
        private readonly TemplateOptions options;
        private readonly CultureInfo culture;

        public TemplateFill(OrdDocument document, TemplateOptions options)
        {
            this.document = document;
            this.options = options;
            culture = options.Culture ?? CultureInfo.CurrentCulture;
        }

        public int Filled { get; private set; }

        public List<MissingValue> Missing { get; } = [];

        /// <summary>Fills the blocks of a container: first the sections that span paragraphs, then what is left.</summary>
        public void Blocks(DocumentContext context, XElement container, TemplateScope scope, int depth)
        {
            WordNames names = context.Names;

            if (depth > MaxSectionDepth)
            {
                return;
            }

            while (TryFindBlockSection(context, container, out XElement start, out XElement end, out Tag open, out Tag close) == true)
            {
                Section(context, container, start, end, open, close, scope, depth, names.Paragraph);
            }

            foreach (XElement child in container.Elements().ToList())
            {
                if (child.Parent is null)
                {
                    continue;
                }

                if (child.Name == names.Paragraph)
                {
                    Paragraph(context, child, scope);
                }
                else if (child.Name == names.Table)
                {
                    Table(context, child, scope, depth);
                }
                else if (child.Name == names.ContentControl)
                {
                    if (TryFillControl(context, child, scope, true) == false && child.Element(names.ContentControlContent) is XElement content)
                    {
                        Blocks(context, content, scope, depth + 1);
                    }
                }
                else if (child.Name == names.W + "customXml" || child.Name == names.Inserted || child.Name == names.MoveTo)
                {
                    Blocks(context, child, scope, depth + 1);
                }
            }
        }

        private static string Location(DocumentContext context, XElement paragraph)
        {
            string text = ParagraphText.Read(context.Names, paragraph).Text;

            if (text.Length > 60)
            {
                text = text[..57] + "...";
            }

            return $"{context.PartName}: \"{text}\"";
        }

        /// <summary>Repeats, keeps or removes what is between the start and end of a section, for each item it has.</summary>
        private void Section(DocumentContext context, XElement container, XElement start, XElement end, Tag open, Tag close, TemplateScope scope, int depth, XName unit)
        {
            WordNames names = context.Names;
            bool found = scope.TryResolve(open.Name, out object? value);
            IReadOnlyList<object?> items = TemplateScope.Items(value);

            if (open.Inverted == true)
            {
                items = [];

                if (TemplateScope.Items(value).Count == 0)
                {
                    items = [null];
                }
            }
            else if (found == false)
            {
                Missing.Add(new MissingValue(open.Name, Location(context, FirstParagraph(names, start))));
            }
            else
            {
                Filled++;
            }

            // The tags go first, so that the copies do not have them.
            RemoveTag(context, end, close);
            RemoveTag(context, start, open);
            List<XElement> region = [];
            bool inside = false;

            foreach (XElement child in container.Elements(unit).ToList())
            {
                if (child == start)
                {
                    inside = true;
                }

                if (inside == true)
                {
                    region.Add(child);
                }

                if (child == end)
                {
                    break;
                }
            }

            // A paragraph that held only a tag is not part of what is repeated.
            if (unit == names.Paragraph)
            {
                if (region.Count > 0 && IsBlank(names, region[0]) == true)
                {
                    region.RemoveAt(0);
                    start.Remove();
                }

                if (region.Count > 0 && region[^1] == end && IsBlank(names, end) == true)
                {
                    region.RemoveAt(region.Count - 1);
                    end.Remove();
                }
            }

            XElement anchor = region.LastOrDefault() ?? start;

            if (anchor.Parent is null)
            {
                anchor = end;
            }

            List<XElement> made = [];

            foreach (object? item in items)
            {
                XElement temporary = new(container.Name, region.Select(element => new XElement(element)));

                if (unit == names.TableRow)
                {
                    foreach (XElement row in temporary.Elements(names.TableRow))
                    {
                        foreach (XElement cell in row.Elements(names.TableCell))
                        {
                            Blocks(context, cell, scope.For(item), depth + 1);
                        }
                    }
                }
                else
                {
                    Blocks(context, temporary, scope.For(item), depth + 1);
                }

                made.AddRange(temporary.Elements().ToList());
            }

            foreach (XElement element in made)
            {
                element.Remove();
            }

            if (anchor.Parent is not null)
            {
                anchor.AddAfterSelf(made);
            }
            else
            {
                container.Add(made);
            }

            foreach (XElement element in region)
            {
                if (document.tracking is not null)
                {
                    document.tracking.DeleteBlock(names, element);
                }
                else
                {
                    element.Remove();
                }
            }

            foreach (XElement element in made)
            {
                document.tracking?.InsertBlock(names, element);
            }

            // A table cell must end with a paragraph.
            if (container.Name == names.TableCell && container.Elements().LastOrDefault()?.Name != names.Paragraph)
            {
                container.Add(new XElement(names.Paragraph));
            }
        }

        private static XElement FirstParagraph(WordNames names, XElement element)
        {
            return element.DescendantsAndSelf(names.Paragraph).FirstOrDefault() ?? element;
        }

        private static bool IsBlank(WordNames names, XElement paragraph)
        {
            return ParagraphText.Read(names, paragraph).Text.Trim().Length == 0
                && paragraph.Descendants().Any(element => element.Name.LocalName is "drawing" or "pict" or "object") == false
                && paragraph.Element(names.ParagraphProperties)?.Element(names.SectionProperties) is null;
        }

        /// <summary>A section whose start tag is in one paragraph of the container and its end tag in a later one.</summary>
        private static bool TryFindBlockSection(DocumentContext context, XElement container, out XElement start, out XElement end, out Tag open, out Tag close)
        {
            WordNames names = context.Names;
            List<XElement> paragraphs = [.. container.Elements(names.Paragraph)];
            start = container;
            end = container;
            open = default;
            close = default;

            for (int index = 0; index < paragraphs.Count; index++)
            {
                List<Tag> tags = Tags(ParagraphText.Read(names, paragraphs[index]).Text);

                foreach (Tag tag in tags.Where(tag => tag.IsOpen == true))
                {
                    if (HasCloseAfter(tags, tag) == true)
                    {
                        continue;
                    }

                    for (int later = index + 1; later < paragraphs.Count; later++)
                    {
                        Tag? closing = Tags(ParagraphText.Read(names, paragraphs[later]).Text).Cast<Tag?>().FirstOrDefault(candidate => candidate!.Value.IsClose == true && candidate.Value.Name == tag.Name);

                        if (closing is not null)
                        {
                            start = paragraphs[index];
                            end = paragraphs[later];
                            open = tag;
                            close = closing.Value;

                            return true;
                        }
                    }
                }
            }

            return false;
        }

        private static bool HasCloseAfter(List<Tag> tags, Tag open)
        {
            return tags.Any(tag => tag.IsClose == true && tag.Name == open.Name && tag.Index > open.Index);
        }

        /// <summary>Removes a tag's text from a paragraph (the start tag's or end tag's paragraph, or a row's).</summary>
        private void RemoveTag(DocumentContext context, XElement block, Tag tag)
        {
            WordNames names = context.Names;

            foreach (XElement paragraph in block.DescendantsAndSelf(names.Paragraph))
            {
                string text = ParagraphText.Read(names, paragraph).Text;
                Tag? found = Tags(text).Cast<Tag?>().FirstOrDefault(candidate => candidate!.Value.Raw == tag.Raw);

                if (found is not null)
                {
                    document.ReplaceIn(context, paragraph, found.Value.Index, found.Value.Index + found.Value.Length, Replacement.Nothing);

                    return;
                }
            }
        }

        /// <summary>Fills a table: first the sections over rows, then each cell.</summary>
        private void Table(DocumentContext context, XElement table, TemplateScope scope, int depth)
        {
            WordNames names = context.Names;

            while (TryFindRowSection(context, table, out XElement start, out XElement end, out Tag open, out Tag close) == true)
            {
                Section(context, table, start, end, open, close, scope, depth, names.TableRow);
            }

            foreach (XElement cell in table.Elements(names.TableRow).SelectMany(row => row.Elements(names.TableCell)).ToList())
            {
                Blocks(context, cell, scope, depth + 1);
            }
        }

        /// <summary>A section whose start tag is in one cell of a row and whose end tag is in another cell or a later row.</summary>
        private static bool TryFindRowSection(DocumentContext context, XElement table, out XElement start, out XElement end, out Tag open, out Tag close)
        {
            WordNames names = context.Names;
            List<XElement> rows = [.. table.Elements(names.TableRow)];
            start = table;
            end = table;
            open = default;
            close = default;

            for (int index = 0; index < rows.Count; index++)
            {
                foreach (XElement cell in rows[index].Elements(names.TableCell))
                {
                    List<Tag> cellTags = [.. cell.Descendants(names.Paragraph).SelectMany(paragraph => Tags(ParagraphText.Read(names, paragraph).Text))];

                    foreach (Tag tag in cellTags.Where(tag => tag.IsOpen == true))
                    {
                        // Opened and closed in the same cell: a section of the cell's own paragraphs.
                        if (cellTags.Any(other => other.IsClose == true && other.Name == tag.Name) == true)
                        {
                            continue;
                        }

                        for (int later = index; later < rows.Count; later++)
                        {
                            IEnumerable<Tag> closing = rows[later].Elements(names.TableCell)
                                .Where(other => other != cell)
                                .SelectMany(other => other.Descendants(names.Paragraph))
                                .SelectMany(paragraph => Tags(ParagraphText.Read(names, paragraph).Text))
                                .Where(candidate => candidate.IsClose == true && candidate.Name == tag.Name);

                            foreach (Tag found in closing)
                            {
                                start = rows[index];
                                end = rows[later];
                                open = tag;
                                close = found;

                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }
    }
}
