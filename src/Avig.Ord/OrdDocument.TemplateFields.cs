using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Avig.Ord.Internal;

namespace Avig.Ord;

public sealed partial class OrdDocument
{
    private enum TagKind
    {
        Value,
        Open,
        Inverted,
        Close,
    }

    /// <summary>A {{…}} tag in a paragraph's text.</summary>
    private readonly record struct Tag(string Name, string Raw, int Index, int Length, TagKind Kind, string? Format)
    {
        public bool IsOpen => Kind is TagKind.Open or TagKind.Inverted;

        public bool IsClose => Kind == TagKind.Close;

        public bool Inverted => Kind == TagKind.Inverted;
    }

    private sealed partial class TemplateFill
    {
        /// <summary>The tags in a text: {{Name}}, {{Name:format}}, {{#Name}}, {{^Name}}, {{/Name}}, with spaces allowed inside the braces.</summary>
        public static List<Tag> Tags(string text)
        {
            List<Tag> tags = [];

            if (text.Contains("{{", StringComparison.Ordinal) == false)
            {
                return tags;
            }

            foreach (Match match in TagPattern().Matches(text))
            {
                string prefix = match.Groups["prefix"].Value;
                TagKind kind = prefix switch
                {
                    "#" => TagKind.Open,
                    "^" => TagKind.Inverted,
                    "/" => TagKind.Close,
                    _ => TagKind.Value,
                };

                string? format = null;

                if (match.Groups["format"].Success == true)
                {
                    format = match.Groups["format"].Value;
                }

                tags.Add(new Tag(match.Groups["name"].Value.Trim(), match.Value, match.Index, match.Length, kind, format));
            }

            return tags;
        }

        /// <summary>
        /// Fills a paragraph: sections that open and close in it (as text, in the formatting of
        /// where they start), then the values, then MERGEFIELD fields and content controls.
        /// </summary>
        private void Paragraph(DocumentContext context, XElement paragraph, TemplateScope scope)
        {
            WordNames names = context.Names;
            string location = Location(context, paragraph);
            string text = ParagraphText.Read(names, paragraph).Text;
            List<(int Start, int End, Replacement Replacement)> changes = [];

            if (text.Contains("{{", StringComparison.Ordinal) == true)
            {
                List<Tag> tags = Tags(text);
                int index = 0;

                while (index < tags.Count)
                {
                    Tag tag = tags[index];

                    if (tag.IsOpen == true)
                    {
                        int close = MatchingClose(tags, index);

                        if (close > index)
                        {
                            string inner = text[(tag.Index + tag.Length)..tags[close].Index];
                            string rendered = Render(inner, tag, scope, location, 0);
                            changes.Add((tag.Index, tags[close].Index + tags[close].Length, Replacement.Text(rendered)));
                            index = close + 1;

                            continue;
                        }

                        // A section tag without its end: it is taken out, and listed.
                        Missing.Add(new MissingValue(tag.Raw, location));
                        changes.Add((tag.Index, tag.Index + tag.Length, Replacement.Nothing));
                        index++;

                        continue;
                    }

                    if (tag.IsClose == true)
                    {
                        Missing.Add(new MissingValue(tag.Raw, location));
                        changes.Add((tag.Index, tag.Index + tag.Length, Replacement.Nothing));
                        index++;

                        continue;
                    }

                    changes.Add((tag.Index, tag.Index + tag.Length, Value(tag, scope, location)));
                    index++;
                }
            }

            // From the end, so that the positions before stay right.
            for (int change = changes.Count - 1; change >= 0; change--)
            {
                (int start, int end, Replacement replacement) = changes[change];

                if (paragraph.Parent is not null)
                {
                    document.ReplaceIn(context, paragraph, start, end, replacement);
                }
            }

            if (paragraph.Parent is null)
            {
                return;
            }

            if (options.MergeFields == true)
            {
                MergeFields(context, paragraph, scope, location);
            }

            if (options.ContentControls == true)
            {
                foreach (XElement control in paragraph.Descendants(names.ContentControl).ToList())
                {
                    TryFillControl(context, control, scope, false);
                }
            }
        }

        /// <summary>The index of the tag that closes the section opened at <paramref name="open"/>, counting sections of the same name inside it.</summary>
        private static int MatchingClose(List<Tag> tags, int open)
        {
            int nested = 0;

            for (int index = open + 1; index < tags.Count; index++)
            {
                if (tags[index].Name != tags[open].Name)
                {
                    continue;
                }

                if (tags[index].IsOpen == true)
                {
                    nested++;
                }
                else if (tags[index].IsClose == true)
                {
                    if (nested == 0)
                    {
                        return index;
                    }

                    nested--;
                }
            }

            return -1;
        }

        /// <summary>A section inside one paragraph, as text: its inner text with its own tags filled, once for each item.</summary>
        private string Render(string inner, Tag open, TemplateScope scope, string location, int depth)
        {
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
                Missing.Add(new MissingValue(open.Name, location));
            }
            else
            {
                Filled++;
            }

            StringBuilder output = new();

            foreach (object? item in items)
            {
                output.Append(RenderText(inner, scope.For(item), location, depth + 1));
            }

            return output.ToString();
        }

        private string RenderText(string text, TemplateScope scope, string location, int depth)
        {
            if (depth > MaxSectionDepth)
            {
                return string.Empty;
            }

            List<Tag> tags = Tags(text);
            StringBuilder output = new();
            int position = 0;
            int index = 0;

            while (index < tags.Count)
            {
                Tag tag = tags[index];
                output.Append(text, position, tag.Index - position);
                position = tag.Index + tag.Length;

                if (tag.IsOpen == true)
                {
                    int close = MatchingClose(tags, index);

                    if (close > index)
                    {
                        output.Append(Render(text[position..tags[close].Index], tag, scope, location, depth));
                        position = tags[close].Index + tags[close].Length;
                        index = close + 1;

                        continue;
                    }

                    Missing.Add(new MissingValue(tag.Raw, location));
                }
                else if (tag.IsClose == true)
                {
                    Missing.Add(new MissingValue(tag.Raw, location));
                }
                else
                {
                    Replacement value = Value(tag, scope, location);

                    if (value.Kind == ReplacementKind.Text)
                    {
                        output.Append(value.TextValue);
                    }
                }

                index++;
            }

            output.Append(text, position, text.Length - position);

            return output.ToString();
        }

        /// <summary>What a value tag becomes: the value as text in its format, or the value itself when it is a <see cref="Replacement"/>; nothing, and listed, when there is none.</summary>
        private Replacement Value(Tag tag, TemplateScope scope, string location)
        {
            if (scope.TryResolve(tag.Name, out object? value) == false)
            {
                Missing.Add(new MissingValue(tag.Name, location));

                return Replacement.Nothing;
            }

            Filled++;

            if (value is Replacement replacement)
            {
                return replacement;
            }

            return Replacement.Text(TemplateScope.Format(value, tag.Format, culture));
        }

        /// <summary>Replaces each MERGEFIELD in the paragraph, simple or complex, with a run of its value in the formatting of its result.</summary>
        private void MergeFields(DocumentContext context, XElement paragraph, TemplateScope scope, string location)
        {
            WordNames names = context.Names;

            foreach (XElement simple in paragraph.Descendants(names.SimpleField).ToList())
            {
                string instruction = simple.Attribute(names.W + "instr")?.Value ?? string.Empty;

                if (IsMergeField(instruction) == false)
                {
                    continue;
                }

                XElement run = MergeRun(names, instruction, simple.Descendants(names.Run).FirstOrDefault()?.Element(names.RunProperties), scope, location);
                Swap(names, [simple], run);
            }

            List<XElement> runs = [.. paragraph.Descendants(names.Run)];

            for (int index = 0; index < runs.Count; index++)
            {
                if (FieldCharType(names, runs[index]) != "begin")
                {
                    continue;
                }

                StringBuilder instruction = new();
                int end = -1;
                int separate = -1;
                int nested = 0;

                for (int later = index; later < runs.Count; later++)
                {
                    string? type = FieldCharType(names, runs[later]);

                    if (later > index && type == "begin")
                    {
                        nested++;
                    }
                    else if (type == "separate" && nested == 0)
                    {
                        separate = later;
                    }
                    else if (type == "end")
                    {
                        if (nested == 0)
                        {
                            end = later;

                            break;
                        }

                        nested--;
                    }

                    if (separate < 0 && nested == 0)
                    {
                        instruction.Append(string.Concat(runs[later].Elements(names.InstructionText).Select(element => element.Value)));
                    }
                }

                if (end < 0)
                {
                    break;
                }

                List<XElement> span = runs[index..(end + 1)];

                if (IsMergeField(instruction.ToString()) == true && span.All(run => run.Parent == runs[index].Parent) == true)
                {
                    XElement? properties = runs[index].Element(names.RunProperties);

                    if (separate >= 0 && separate + 1 < end)
                    {
                        properties = runs[separate + 1].Element(names.RunProperties) ?? properties;
                    }

                    Swap(names, span, MergeRun(names, instruction.ToString(), properties, scope, location));
                }

                index = end;
            }
        }

        private static string? FieldCharType(WordNames names, XElement run)
        {
            return run.Element(names.FieldChar)?.Attribute(names.W + "fldCharType")?.Value;
        }

        private static bool IsMergeField(string instruction)
        {
            return instruction.TrimStart().StartsWith("MERGEFIELD", StringComparison.OrdinalIgnoreCase) == true;
        }

        /// <summary>The run a merge field becomes, with Word's switches: \@ and \# give the format, \* Upper, Lower, Caps and FirstCap the case, \b and \f text before and after a value that is not empty.</summary>
        private XElement MergeRun(WordNames names, string instruction, XElement? properties, TemplateScope scope, string location)
        {
            string name = FieldInstruction.FirstArgument(instruction);
            string text = string.Empty;

            if (scope.TryResolve(name, out object? value) == true)
            {
                Filled++;
                string? format = Switch(instruction, "@") ?? Switch(instruction, "#");
                text = TemplateScope.Format(value, format, culture);
                string? casing = Switch(instruction, "*");

                text = casing?.ToUpperInvariant() switch
                {
                    "UPPER" => text.ToUpper(culture),
                    "LOWER" => text.ToLower(culture),
                    "CAPS" => culture.TextInfo.ToTitleCase(text.ToLower(culture)),
                    "FIRSTCAP" when text.Length > 0 => char.ToUpper(text[0], culture) + text[1..],
                    _ => text,
                };

                if (text.Length > 0)
                {
                    text = (Switch(instruction, "b") ?? string.Empty) + text + (Switch(instruction, "f") ?? string.Empty);
                }
            }
            else
            {
                Missing.Add(new MissingValue(name, location));
            }

            XElement run = new(names.Run);

            if (properties is not null)
            {
                run.Add(new XElement(properties));
            }

            ContentXml.AddText(names, run, text);

            return run;
        }

        /// <summary>The argument of a field switch such as \@ "yyyy-MM-dd", unquoted; null when the instruction does not have it.</summary>
        private static string? Switch(string instruction, string name)
        {
            Match match = Regex.Match(instruction, @"\\" + Regex.Escape(name) + @"\s+(?:""(?<value>[^""]*)""|(?<value>\S+))", RegexOptions.None, TimeSpan.FromSeconds(1));
            string? value = null;

            if (match.Success == true)
            {
                value = match.Groups["value"].Value;
            }

            return value;
        }

        /// <summary>Puts <paramref name="run"/> where <paramref name="span"/> was: straight, or marked as a tracked change.</summary>
        private void Swap(WordNames names, List<XElement> span, XElement run)
        {
            span[0].AddBeforeSelf(run);

            if (document.tracking is not null)
            {
                List<XElement> spanRuns = [.. span.SelectMany(element => element.DescendantsAndSelf(names.Run))];
                document.tracking.Delete(names, spanRuns);
                document.tracking.Insert(names, [run]);

                foreach (XElement simple in span.Where(element => element.Name == names.SimpleField && element.Parent is not null))
                {
                    simple.ReplaceWith(simple.Nodes());
                }

                return;
            }

            foreach (XElement element in span)
            {
                element.Remove();
            }
        }

        /// <summary>Fills a content control whose tag (or title) has a value; false when it has none, so that its content is filled as ordinary content.</summary>
        private bool TryFillControl(DocumentContext context, XElement control, TemplateScope scope, bool block)
        {
            if (options.ContentControls == false)
            {
                return false;
            }

            WordNames names = context.Names;
            XElement? properties = control.Element(names.ContentControlProperties);
            string? name = names.ValueOf(properties, "tag") ?? names.ValueOf(properties, "alias");

            if (name is null || name.Length == 0)
            {
                return false;
            }

            XElement? content = control.Element(names.ContentControlContent);

            if (scope.TryResolve(name, out object? value) == false)
            {
                // A control that is not about the data (a form field) keeps what it has.
                return false;
            }

            Filled++;

            if (content is null)
            {
                content = new XElement(names.ContentControlContent);
                control.Add(content);
            }

            properties?.Element(names.W + "showingPlcHdr")?.Remove();
            string text = TemplateScope.Format(value, null, culture);

            if (value is Replacement replacement && replacement.Kind == ReplacementKind.Text)
            {
                text = replacement.TextValue;
            }

            XElement? runProperties = content.Descendants(names.Run).FirstOrDefault()?.Element(names.RunProperties);
            XElement run = new(names.Run);

            if (runProperties is not null)
            {
                run.Add(new XElement(runProperties));
            }

            ContentXml.AddText(names, run, text);

            if (block == true)
            {
                XElement? paragraphProperties = content.Descendants(names.ParagraphProperties).FirstOrDefault();
                XElement paragraph = new(names.Paragraph);

                if (paragraphProperties is not null)
                {
                    paragraph.Add(new XElement(paragraphProperties));
                }

                paragraph.Add(run);
                Replace(names, content, paragraph);
            }
            else
            {
                Replace(names, content, run);
            }

            return true;
        }

        private void Replace(WordNames names, XElement content, XElement replacement)
        {
            if (document.tracking is not null)
            {
                foreach (XElement block in content.Elements().ToList())
                {
                    if (block.Name == names.Run)
                    {
                        document.tracking.Delete(names, [block]);
                    }
                    else
                    {
                        document.tracking.DeleteBlock(names, block);
                    }
                }

                content.Add(replacement);

                if (replacement.Name == names.Run)
                {
                    document.tracking.Insert(names, [replacement]);
                }
                else
                {
                    document.tracking.InsertBlock(names, replacement);
                }

                return;
            }

            content.RemoveNodes();
            content.Add(replacement);
        }

        [GeneratedRegex(@"\{\{\s*(?<prefix>[#^/]?)\s*(?<name>[^{}:#^/\s][^{}:]*?|\.)\s*(?::(?<format>[^{}]*))?\s*\}\}", RegexOptions.CultureInvariant, 1000)]
        private static partial Regex TagPattern();
    }
}
