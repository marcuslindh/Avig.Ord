using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace Avig.Ord.Internal;

/// <summary>
/// Reads what is in a paragraph into <see cref="Inline"/>s, in the document's revision view:
/// runs (with tabs and line breaks in their text), breaks, hyperlinks, fields (the result,
/// with the instruction kept apart), pictures, note references, text boxes and content controls.
/// </summary>
internal static class InlineReader
{
    /// <summary>How deep complex fields may nest. Word nests a few levels (an IF with a MERGEFIELD in it); 32 is far past that. EventReader stops at the same depth.</summary>
    internal const int MaxFieldNesting = 32;

    public static List<Inline> Read(DocumentContext context, XElement paragraph, RunFormat inherited)
    {
        State state = new(context, inherited);
        Walk(state, paragraph, RevisionKind.None);
        state.CloseOpenFields();

        return state.Root;
    }

    /// <summary>Every inline, with those inside hyperlinks, fields and content controls in their place.</summary>
    public static IEnumerable<Inline> Flatten(IEnumerable<Inline> inlines)
    {
        foreach (Inline inline in inlines)
        {
            yield return inline;

            IReadOnlyList<Inline>? children = inline switch
            {
                Hyperlink link => link.Inlines,
                Field field => field.Result,
                ContentControl control => control.Inlines,
                _ => null,
            };

            if (children is null)
            {
                continue;
            }

            foreach (Inline child in Flatten(children))
            {
                yield return child;
            }
        }
    }

    private static void Walk(State state, XElement container, RevisionKind revision)
    {
        WordNames names = state.Context.Names;

        foreach (XElement child in container.Elements())
        {
            XName name = child.Name;

            if (name == names.Run)
            {
                ReadRun(state, child, revision);
            }
            else if (name == names.Hyperlink)
            {
                List<Inline> link = state.Begin();
                Walk(state, child, revision);
                state.Add(new Hyperlink(state.Context, child, state.End(link)));
            }
            else if (name == names.SimpleField)
            {
                List<Inline> result = state.Begin();
                Walk(state, child, revision);
                state.Add(new Field(state.Context, child, child.Attribute(names.W + "instr")?.Value ?? string.Empty, state.End(result)));
            }
            else if (name == names.ContentControl)
            {
                List<Inline> controlled = state.Begin();
                XElement? content = child.Element(names.ContentControlContent);

                if (content is not null)
                {
                    Walk(state, content, revision);
                }

                state.Add(new ContentControl(state.Context, child, state.End(controlled)));
            }
            else if (name == names.Inserted || name == names.MoveTo)
            {
                if (state.Context.View != RevisionView.Original)
                {
                    Walk(state, child, Marked(state, RevisionKind.Inserted));
                }
            }
            else if (name == names.Deleted || name == names.MoveFrom)
            {
                if (state.Context.View != RevisionView.Accepted)
                {
                    Walk(state, child, Marked(state, RevisionKind.Deleted));
                }
            }
            else if (name == names.W + "smartTag" || name == names.W + "customXml" || name == names.W + "bdo" || name == names.W + "dir")
            {
                Walk(state, child, revision);
            }
            else if (name.LocalName == "AlternateContent")
            {
                XElement? chosen = Choose(child);

                if (chosen is not null)
                {
                    Walk(state, chosen, revision);
                }
            }
        }
    }

    private static RevisionKind Marked(State state, RevisionKind kind)
    {
        if (state.Context.View == RevisionView.Marked)
        {
            return kind;
        }

        return RevisionKind.None;
    }

    /// <summary>mc:Choice when it holds DrawingML (what this package reads), otherwise mc:Fallback.</summary>
    private static XElement? Choose(XElement alternate)
    {
        XElement? choice = alternate.Elements().FirstOrDefault(element => element.Name.LocalName == "Choice");
        XElement? fallback = alternate.Elements().FirstOrDefault(element => element.Name.LocalName == "Fallback");

        if (choice is not null && choice.Descendants().Any(element => element.Name.LocalName == "drawing") == true)
        {
            return choice;
        }

        return fallback ?? choice;
    }

    private static void ReadRun(State state, XElement run, RevisionKind revision)
    {
        WordNames names = state.Context.Names;
        StringBuilder text = new();

        foreach (XElement child in run.Elements())
        {
            string local = child.Name.LocalName;

            if (child.Name.Namespace != names.W && local != "AlternateContent")
            {
                continue;
            }

            switch (local)
            {
                case "fldChar":
                    state.Text(run, text, revision);
                    state.FieldChar(run, child.Attribute(names.W + "fldCharType")?.Value);

                    break;

                case "instrText":
                    state.Instruction(child.Value);

                    break;

                case "delInstrText":
                    if (revision == RevisionKind.Deleted || state.Context.View == RevisionView.Original)
                    {
                        state.Instruction(child.Value);
                    }

                    break;

                case "t":
                    text.Append(child.Value);

                    break;

                case "delText":
                    // Deleted text is only in a run inside w:del, which the Accepted view never walks into.
                    text.Append(child.Value);

                    break;

                case "tab":
                case "ptab":
                    text.Append('\t');

                    break;

                case "cr":
                    text.Append('\n');

                    break;

                case "br":
                    string? type = child.Attribute(names.Type)?.Value;

                    if (type is "page" or "column")
                    {
                        state.Text(run, text, revision);
                        BreakKind kind = BreakKind.Page;

                        if (type == "column")
                        {
                            kind = BreakKind.Column;
                        }

                        state.Add(new Break(state.Context, child, kind));
                    }
                    else
                    {
                        text.Append('\n');
                    }

                    break;

                case "noBreakHyphen":
                    text.Append('‑');

                    break;

                case "sym":
                    text.Append(Symbol(names, child));

                    break;

                case "footnoteReference":
                case "endnoteReference":
                    state.Text(run, text, revision);
                    NoteKind note = NoteKind.Footnote;

                    if (local == "endnoteReference")
                    {
                        note = NoteKind.Endnote;
                    }

                    state.Add(new NoteReference(state.Context, child, note, child.Attribute(names.WId)?.Value ?? string.Empty));

                    break;

                case "drawing":
                case "pict":
                case "object":
                    state.Text(run, text, revision);
                    ReadDrawing(state, child);

                    break;

                case "AlternateContent":
                    state.Text(run, text, revision);
                    XElement? chosen = Choose(child);

                    if (chosen is not null)
                    {
                        foreach (XElement drawing in chosen.Elements().Where(element => element.Name.LocalName is "drawing" or "pict"))
                        {
                            ReadDrawing(state, drawing);
                        }
                    }

                    break;

                default:
                    break;
            }
        }

        state.Text(run, text, revision);
    }

    private static string Symbol(WordNames names, XElement symbol)
    {
        return Symbol(symbol.Attribute(names.W + "char")?.Value);
    }

    /// <summary>The character a w:sym stands for, from its hexadecimal w:char; empty when that is no character.</summary>
    internal static string Symbol(string? code)
    {
        string text = string.Empty;

        if (code is not null && int.TryParse(code, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int value) == true && value is > 0 and < 0x110000 and not (>= 0xD800 and <= 0xDFFF))
        {
            text = char.ConvertFromUtf32(value);
        }

        return text;
    }

    /// <summary>A picture (DrawingML or VML) or a text box; charts, shapes and embedded objects without either give nothing.</summary>
    private static void ReadDrawing(State state, XElement drawing)
    {
        DocumentContext context = state.Context;

        // What is inside a text box belongs to the box's own paragraphs: a text box or picture
        // there is read when they are, in their revision view and table rows. Looking into the
        // boxes too gave a nested box's text twice, and the text of rows the view leaves out.
        List<XElement> outside = OutsideTextBoxes(drawing);

        foreach (XElement textBox in outside.Where(element => element.Name.LocalName == "txbxContent").Take(16))
        {
            state.Add(new TextBox(context, textBox, Block.Read(context, textBox)));
        }

        XElement? blip = outside.FirstOrDefault(element => element.Name.LocalName == "blip");

        if (blip is not null)
        {
            XElement? extent = outside.FirstOrDefault(element => element.Name.LocalName == "extent");
            XElement? properties = outside.FirstOrDefault(element => element.Name.LocalName == "docPr");
            XElement? svg = blip.Descendants().FirstOrDefault(element => element.Name.LocalName == "svgBlip");

            state.Add(new Picture(
                context,
                drawing,
                RelationshipAttribute(blip, "embed") ?? RelationshipAttribute(blip, "link"),
                RelationshipAttribute(svg, "embed"),
                FormatReader.Emu(extent?.Attribute("cx")?.Value) ?? Length.Zero,
                FormatReader.Emu(extent?.Attribute("cy")?.Value) ?? Length.Zero,
                properties?.Attribute("descr")?.Value ?? string.Empty,
                properties?.Attribute("title")?.Value ?? string.Empty));

            return;
        }

        XElement? image = outside.FirstOrDefault(element => element.Name.LocalName == "imagedata");

        if (image is not null)
        {
            state.Add(new Picture(context, drawing, RelationshipAttribute(image, "id"), null, Length.Zero, Length.Zero, image.Attribute(XName.Get("title", "urn:schemas-microsoft-com:office:office"))?.Value ?? string.Empty, string.Empty));
        }
    }

    /// <summary>The elements under a drawing in document order, without going into a text box: a w:txbxContent is in the list, what it holds is not.</summary>
    private static List<XElement> OutsideTextBoxes(XElement drawing)
    {
        List<XElement> found = [];
        Stack<XElement> pending = new();

        foreach (XElement child in drawing.Elements().Reverse())
        {
            pending.Push(child);
        }

        while (pending.Count > 0)
        {
            XElement element = pending.Pop();
            found.Add(element);

            if (element.Name.LocalName == "txbxContent")
            {
                continue;
            }

            foreach (XElement child in element.Elements().Reverse())
            {
                pending.Push(child);
            }
        }

        return found;
    }

    private static string? RelationshipAttribute(XElement? element, string localName)
    {
        return element?.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == localName && attribute.Name.NamespaceName.Contains("relationships", StringComparison.Ordinal) == true)?.Value;
    }

    /// <summary>
    /// Where inlines go as the walk goes on: the paragraph's list, or the list of the hyperlink,
    /// content control or field result being read. Complex fields (w:fldChar begin, separate,
    /// end) can span runs, hyperlinks and even nest, so they are a stack of their own.
    /// </summary>
    private sealed class State
    {
        private readonly Stack<List<Inline>> targets = new();
        private readonly Stack<OpenField> fields = new();
        private readonly HashSet<List<Inline>> closed = new(ReferenceEqualityComparer.Instance);
        private int ignoredFields;
        private readonly RunFormat inherited;

        public State(DocumentContext context, RunFormat inherited)
        {
            Context = context;
            this.inherited = inherited;
            targets.Push(Root);
        }

        public DocumentContext Context { get; }

        public List<Inline> Root { get; } = [];

        /// <summary>Starts the list of a container (a hyperlink, a simple field, a content control); <see cref="End"/> ends it.</summary>
        public List<Inline> Begin()
        {
            List<Inline> list = [];
            targets.Push(list);

            return list;
        }

        public List<Inline> End(List<Inline> list)
        {
            // Only a broken file can get here without the list on the stack; the root must stay.
            if (targets.Contains(list) == false)
            {
                return list;
            }

            // A field begun inside the container and not ended there has its result above the
            // container's list; it is closed with the container, so that its inlines are not lost.
            while (targets.Count > 1 && ReferenceEquals(targets.Peek(), list) == false)
            {
                if (fields.Count > 0 && ReferenceEquals(targets.Peek(), fields.Peek().Result) == true)
                {
                    Close();
                }
                else
                {
                    closed.Remove(targets.Pop());
                }
            }

            if (targets.Count > 1)
            {
                targets.Pop();
            }

            PopClosedResults();

            return list;
        }

        public void Add(Inline inline)
        {
            if (fields.Count > 0 && fields.Peek().InInstruction == true)
            {
                return;
            }

            targets.Peek().Add(inline);
        }

        /// <summary>Adds the text gathered so far as a run, and empties the buffer.</summary>
        public void Text(XElement run, StringBuilder text, RevisionKind revision)
        {
            if (text.Length == 0)
            {
                return;
            }

            Add(new Run(Context, run, text.ToString(), revision, inherited));
            text.Clear();
        }

        public void Instruction(string text)
        {
            if (fields.Count > 0 && fields.Peek().InInstruction == true)
            {
                fields.Peek().Instruction.Append(text);
            }
        }

        public void FieldChar(XElement run, string? type)
        {
            // Fields nested deeper than MaxFieldNesting are read as part of the field they are in:
            // their begin, separate and end are passed over. Twenty thousand fields begun in a row,
            // each in the result of the one before, nested Field objects that deep, and reading
            // their text ran out of stack; the XML itself is flat.
            if (ignoredFields > 0 || (type == "begin" && fields.Count >= MaxFieldNesting))
            {
                if (type == "begin")
                {
                    ignoredFields++;
                }
                else if (type == "end")
                {
                    ignoredFields--;
                }

                return;
            }

            switch (type)
            {
                case "begin":
                    // Nested inside another field's instruction (as in an IF field) the inner
                    // field's result is part of the outer instruction; keep it out of the text.
                    fields.Push(new OpenField(run, targets.Peek()));
                    targets.Push(fields.Peek().Result);

                    break;

                case "separate":
                    if (fields.Count > 0)
                    {
                        fields.Peek().InInstruction = false;
                    }

                    break;

                case "end":
                    if (fields.Count > 0)
                    {
                        Close();
                    }

                    break;

                default:
                    break;
            }
        }

        public void CloseOpenFields()
        {
            while (fields.Count > 0)
            {
                Close();
            }
        }

        private void Close()
        {
            OpenField field = fields.Pop();

            // A field that began before a container (a hyperlink, a content control) and ends
            // inside it leaves the container open: what is left of the container still goes into
            // the container, which is inside the field's result, and the result is taken off the
            // stack when the container ends. Popping the container here put the rest of a
            // hyperlink before its start, and then took the paragraph's own list off the stack.
            if (ReferenceEquals(targets.Peek(), field.Result) == true)
            {
                targets.Pop();
                PopClosedResults();
            }
            else if (targets.Contains(field.Result) == true)
            {
                closed.Add(field.Result);
            }

            Field made = new(Context, field.Start, field.Instruction.ToString(), field.Result);
            bool insideInstruction = fields.Count > 0 && fields.Peek().InInstruction == true;

            if (insideInstruction == false)
            {
                field.Parent.Add(made);
            }
        }

        /// <summary>Takes the results of fields that have ended off the top of the stack, once the containers above them have ended.</summary>
        private void PopClosedResults()
        {
            while (targets.Count > 1 && closed.Remove(targets.Peek()) == true)
            {
                targets.Pop();
            }
        }

        private sealed class OpenField
        {
            public OpenField(XElement start, List<Inline> parent)
            {
                Start = start;
                Parent = parent;
            }

            public XElement Start { get; }

            public List<Inline> Parent { get; }

            public StringBuilder Instruction { get; } = new();

            public List<Inline> Result { get; } = [];

            public bool InInstruction { get; set; } = true;
        }
    }
}
