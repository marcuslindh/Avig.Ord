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
                state.Begin();
                Walk(state, child, revision);
                state.Add(new Hyperlink(state.Context, child, state.End()));
            }
            else if (name == names.SimpleField)
            {
                state.Begin();
                Walk(state, child, revision);
                state.Add(new Field(state.Context, child, child.Attribute(names.W + "instr")?.Value ?? string.Empty, state.End()));
            }
            else if (name == names.ContentControl)
            {
                state.Begin();
                XElement? content = child.Element(names.ContentControlContent);

                if (content is not null)
                {
                    Walk(state, content, revision);
                }

                state.Add(new ContentControl(state.Context, child, state.End()));
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
        string? code = symbol.Attribute(names.W + "char")?.Value;
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

        foreach (XElement textBox in drawing.Descendants().Where(element => element.Name.LocalName == "txbxContent").Take(16))
        {
            state.Add(new TextBox(context, textBox, Block.Read(context, textBox)));
        }

        XElement? blip = drawing.Descendants().FirstOrDefault(element => element.Name.LocalName == "blip");

        if (blip is not null)
        {
            XElement? extent = drawing.Descendants().FirstOrDefault(element => element.Name.LocalName == "extent");
            XElement? properties = drawing.Descendants().FirstOrDefault(element => element.Name.LocalName == "docPr");
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

        XElement? image = drawing.Descendants().FirstOrDefault(element => element.Name.LocalName == "imagedata");

        if (image is not null)
        {
            state.Add(new Picture(context, drawing, RelationshipAttribute(image, "id"), null, Length.Zero, Length.Zero, image.Attribute(XName.Get("title", "urn:schemas-microsoft-com:office:office"))?.Value ?? string.Empty, string.Empty));
        }
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
        private readonly RunFormat inherited;

        public State(DocumentContext context, RunFormat inherited)
        {
            Context = context;
            this.inherited = inherited;
            targets.Push(Root);
        }

        public DocumentContext Context { get; }

        public List<Inline> Root { get; } = [];

        public void Begin()
        {
            targets.Push([]);
        }

        public List<Inline> End()
        {
            List<Inline> list = targets.Pop();

            // A field begun inside the container and not ended there writes into this list; it
            // is closed with it, so that its inlines are not lost.
            while (fields.Count > 0 && ReferenceEquals(fields.Peek().Parent, list) == true)
            {
                Close();
            }

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

            // Pop the field's result list, and anything left open inside it.
            while (targets.Count > 1 && ReferenceEquals(targets.Peek(), field.Result) == false)
            {
                targets.Pop();
            }

            if (targets.Count > 1)
            {
                targets.Pop();
            }

            Field made = new(Context, field.Start, field.Instruction.ToString(), field.Result);
            bool insideInstruction = fields.Count > 0 && fields.Peek().InInstruction == true;

            if (insideInstruction == false)
            {
                field.Parent.Add(made);
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
