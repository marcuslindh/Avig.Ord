using System.Runtime.CompilerServices;
using System.Xml;
using System.Xml.Linq;
using Avig.OpenXml;
using Avig.Ord.Internal;

namespace Avig.Ord;

/// <summary>What a <see cref="DocumentEvent"/> is.</summary>
public enum DocumentEventKind
{
    /// <summary>A paragraph starts.</summary>
    ParagraphStart,

    /// <summary>A paragraph ends.</summary>
    ParagraphEnd,

    /// <summary>Text, in <see cref="DocumentEvent.Text"/>: what a w:t holds, '\t' for a tab, '\n' for a line break.</summary>
    Text,

    /// <summary>A page or column break.</summary>
    Break,

    /// <summary>A table starts.</summary>
    TableStart,

    /// <summary>A table ends.</summary>
    TableEnd,

    /// <summary>A table row starts.</summary>
    RowStart,

    /// <summary>A table row ends.</summary>
    RowEnd,

    /// <summary>A table cell starts.</summary>
    CellStart,

    /// <summary>A table cell ends.</summary>
    CellEnd,
}

/// <summary>One step of <see cref="DocxReader.ReadEventsAsync"/>.</summary>
/// <param name="Kind">What happened.</param>
/// <param name="Text">The text, for <see cref="DocumentEventKind.Text"/>; empty otherwise.</param>
/// <param name="TableDepth">How many tables the event is inside: 0 in the body, 1 in a table, 2 in a table in a table.</param>
public readonly record struct DocumentEvent(DocumentEventKind Kind, string Text, int TableDepth);

/// <summary>Reads the body as <see cref="DocumentEvent"/>s straight from the XML, without making an element tree.</summary>
internal static class EventReader
{
    public static async IAsyncEnumerable<DocumentEvent> ReadAsync(XmlReader reader, RevisionView view, [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        int bodyDepth = reader.Depth;
        int tables = 0;

        // The depth of the element being skipped (a deletion in the Accepted view, a text box, a
        // field instruction …), or -1; and how deep fields are nested, with whether each is still
        // in its instruction.
        int skipUntil = -1;
        int pendingRow = -1;
        Stack<(bool InInstruction, int Container)> fields = new();
        bool inInstruction = false;

        // How many hyperlinks, simple fields and content controls the reader is in: a field begun
        // inside one ends with it, as Paragraph.Inlines reads it.
        int containers = 0;
        int ignoredFields = 0;
        bool positioned = false;
        bool first = true;

        XmlReader source = reader;
        XmlReader? paragraph = null;
        bool more = await source.ReadAsync();

        while (more == true)
        {
            cancellationToken.ThrowIfCancellationRequested();

            // ReadElementContentAsString leaves the reader on the node after the element; it must not be read past.
            if (positioned == true)
            {
                positioned = false;
            }
            else if (first == false)
            {
                more = await source.ReadAsync();

                if (more == false && paragraph is not null)
                {
                    // The end of a paragraph read from memory: back to the part, which is already
                    // on the node after the paragraph.
                    paragraph.Dispose();
                    paragraph = null;
                    source = reader;
                    more = true;
                    positioned = true;

                    continue;
                }

                if (more == false)
                {
                    break;
                }
            }

            first = false;

            if (paragraph is null && OpcXml.IsEnd(source, bodyDepth) == true)
            {
                yield break;
            }

            if (skipUntil >= 0)
            {
                if (source.NodeType == XmlNodeType.EndElement && source.Depth == skipUntil)
                {
                    skipUntil = -1;
                }

                continue;
            }

            // A row marked deleted (Accepted view) or inserted (Original view) in its w:trPr is
            // left out, as Table.Rows leaves it out; the properties come before the cells.
            if (pendingRow >= 0 && source.NodeType is XmlNodeType.Element or XmlNodeType.EndElement)
            {
                int row = pendingRow;

                if (source.NodeType == XmlNodeType.Element && source.Depth == row + 1 && source.LocalName is "tblPrEx" or "trPr")
                {
                    XNode properties = await XNode.ReadFromAsync(source, cancellationToken);
                    positioned = true;

                    if (properties is XElement element && element.Name.LocalName == "trPr")
                    {
                        pendingRow = -1;

                        if (IsHidden(element, view) == true)
                        {
                            skipUntil = row;

                            continue;
                        }

                        yield return new DocumentEvent(DocumentEventKind.RowStart, string.Empty, tables);
                    }

                    continue;
                }

                pendingRow = -1;
                yield return new DocumentEvent(DocumentEventKind.RowStart, string.Empty, tables);
            }

            if (source.NodeType == XmlNodeType.Element)
            {
                string name = source.LocalName;
                bool empty = source.IsEmptyElement;

                if (Skips(name, view) == true)
                {
                    if (empty == false)
                    {
                        skipUntil = source.Depth;
                    }

                    continue;
                }

                // Only WordprocessingML: the text of an equation (m:t) is no more in the body's
                // text here than in Paragraph.Text.
                if (source.NamespaceURI is not (WordNames.TransitionalMain or WordNames.StrictMain))
                {
                    continue;
                }

                switch (name)
                {
                    case "p":
                        // A paragraph is read into memory first, to tell whether the view has it
                        // at all (Block.IsGoneInView, as the blocks have it), and then read from there.
                        if (empty == false && paragraph is null)
                        {
                            XElement element = (XElement)await XNode.ReadFromAsync(source, cancellationToken);

                            // Left on the node after the paragraph, whether it is read or not.
                            positioned = true;

                            if (Block.IsGoneInView(WordNames.For(element), view, element) == true)
                            {
                                continue;
                            }

                            // XElement.CreateReader cannot read async, so the paragraph is read back from its text.
                            XmlReaderSettings settings = new()
                            {
                                Async = true,
                                DtdProcessing = DtdProcessing.Prohibit,
                                XmlResolver = null,
                            };

                            paragraph = XmlReader.Create(new StringReader(element.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)), settings);
                            source = paragraph;
                            await source.ReadAsync();
                            positioned = true;

                            continue;
                        }

                        // Fields are read a paragraph at a time, as Paragraph.Inlines reads them: one left
                        // open in a paragraph before does not hide this one's text as its instruction.
                        fields.Clear();
                        ignoredFields = 0;
                        inInstruction = false;
                        yield return new DocumentEvent(DocumentEventKind.ParagraphStart, string.Empty, tables);

                        if (empty == true)
                        {
                            yield return new DocumentEvent(DocumentEventKind.ParagraphEnd, string.Empty, tables);
                        }

                        break;

                    case "tbl":
                        yield return new DocumentEvent(DocumentEventKind.TableStart, string.Empty, tables);

                        // An empty element has no end element to give the end event on.
                        if (empty == true)
                        {
                            yield return new DocumentEvent(DocumentEventKind.TableEnd, string.Empty, tables);
                        }
                        else
                        {
                            tables++;
                        }

                        break;

                    case "tr":
                        if (empty == true)
                        {
                            yield return new DocumentEvent(DocumentEventKind.RowStart, string.Empty, tables);
                            yield return new DocumentEvent(DocumentEventKind.RowEnd, string.Empty, tables);
                        }
                        else
                        {
                            // The row starts once its properties have said whether the view has it.
                            pendingRow = source.Depth;
                        }

                        break;

                    case "tc":
                        yield return new DocumentEvent(DocumentEventKind.CellStart, string.Empty, tables);

                        if (empty == true)
                        {
                            yield return new DocumentEvent(DocumentEventKind.CellEnd, string.Empty, tables);
                        }

                        break;

                    case "hyperlink":
                    case "fldSimple":
                    case "sdt":
                        if (empty == false)
                        {
                            containers++;
                        }

                        break;

                    case "noBreakHyphen":
                        if (inInstruction == false)
                        {
                            yield return new DocumentEvent(DocumentEventKind.Text, "‑", tables);
                        }

                        break;

                    case "sym":
                        OpcXml.TryGetAttribute(source, "char", out string? code);
                        string symbol = InlineReader.Symbol(code);

                        if (inInstruction == false && symbol.Length > 0)
                        {
                            yield return new DocumentEvent(DocumentEventKind.Text, symbol, tables);
                        }

                        break;

                    case "fldChar":
                        OpcXml.TryGetAttribute(source, "fldCharType", out string? type);

                        // As Paragraph.Inlines: fields nested past InlineReader.MaxFieldNesting are part of the one they are in.
                        if (ignoredFields > 0 || (type == "begin" && fields.Count >= InlineReader.MaxFieldNesting))
                        {
                            if (type == "begin")
                            {
                                ignoredFields++;
                            }
                            else if (type == "end")
                            {
                                ignoredFields--;
                            }
                        }
                        else if (type == "begin")
                        {
                            fields.Push((true, containers));
                        }
                        else if (type == "separate" && fields.Count > 0)
                        {
                            (bool _, int container) = fields.Pop();
                            fields.Push((false, container));
                        }
                        else if (type == "end" && fields.Count > 0)
                        {
                            fields.Pop();
                        }

                        inInstruction = fields.Any(field => field.InInstruction == true);

                        break;

                    case "t":
                    case "delText":
                        string text = await source.ReadElementContentAsStringAsync();
                        positioned = true;

                        if (inInstruction == false && text.Length > 0)
                        {
                            yield return new DocumentEvent(DocumentEventKind.Text, text, tables);
                        }

                        break;

                    case "tab":
                    case "ptab":
                        if (inInstruction == false)
                        {
                            yield return new DocumentEvent(DocumentEventKind.Text, "\t", tables);
                        }

                        break;

                    case "cr":
                        if (inInstruction == false)
                        {
                            yield return new DocumentEvent(DocumentEventKind.Text, "\n", tables);
                        }

                        break;

                    case "br":
                        if (inInstruction == false)
                        {
                            OpcXml.TryGetAttribute(source, "type", out string? breakType);

                            if (breakType is "page" or "column")
                            {
                                yield return new DocumentEvent(DocumentEventKind.Break, string.Empty, tables);
                            }
                            else
                            {
                                yield return new DocumentEvent(DocumentEventKind.Text, "\n", tables);
                            }
                        }

                        break;

                    default:
                        break;
                }

                continue;
            }

            if (source.NodeType != XmlNodeType.EndElement || source.NamespaceURI is not (WordNames.TransitionalMain or WordNames.StrictMain))
            {
                continue;
            }

            switch (source.LocalName)
            {
                case "p":
                    fields.Clear();
                    ignoredFields = 0;
                    inInstruction = false;
                    yield return new DocumentEvent(DocumentEventKind.ParagraphEnd, string.Empty, tables);

                    break;

                case "hyperlink":
                case "fldSimple":
                case "sdt":
                    containers--;

                    while (fields.Count > 0 && fields.Peek().Container > containers)
                    {
                        fields.Pop();
                    }

                    inInstruction = fields.Any(field => field.InInstruction == true);

                    break;

                case "tbl":
                    tables--;
                    yield return new DocumentEvent(DocumentEventKind.TableEnd, string.Empty, tables);

                    break;

                case "tr":
                    yield return new DocumentEvent(DocumentEventKind.RowEnd, string.Empty, tables);

                    break;

                case "tc":
                    yield return new DocumentEvent(DocumentEventKind.CellEnd, string.Empty, tables);

                    break;

                default:
                    break;
            }
        }
    }

    /// <summary>A row the view does not have: marked deleted in the Accepted view, inserted in the Original view, as <see cref="TableRow.IsVisible"/> decides.</summary>
    private static bool IsHidden(XElement properties, RevisionView view)
    {
        bool deleted = properties.Elements().Any(element => element.Name.LocalName == "del");
        bool inserted = properties.Elements().Any(element => element.Name.LocalName == "ins");

        return (deleted == true && view == RevisionView.Accepted) || (inserted == true && view == RevisionView.Original);
    }

    /// <summary>
    /// Elements whose content is not part of the body's text in the view: deletions (Accepted),
    /// insertions (Original), field instructions, properties, drawings and text boxes (as in a
    /// paragraph's <see cref="Paragraph.Text"/>), and the fallback of alternate content.
    /// </summary>
    private static bool Skips(string name, RevisionView view)
    {
        return name switch
        {
            "del" or "moveFrom" => view == RevisionView.Accepted,
            "ins" or "moveTo" => view == RevisionView.Original,
            "instrText" or "delInstrText" or "pPr" or "rPr" or "tblPr" or "trPr" or "tcPr" or "tblGrid" or "sectPr" or "sdtPr" or "sdtEndPr" => true,
            "drawing" or "pict" or "object" or "Fallback" or "txbxContent" => true,
            _ => false,
        };
    }
}
