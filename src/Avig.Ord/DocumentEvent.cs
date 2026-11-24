using System.Runtime.CompilerServices;
using System.Xml;
using Avig.OpenXml;

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
        Stack<bool> fields = new();
        bool inInstruction = false;
        bool positioned = false;
        bool first = true;

        bool more = await reader.ReadAsync();

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
                more = await reader.ReadAsync();

                if (more == false)
                {
                    break;
                }
            }

            first = false;

            if (OpcXml.IsEnd(reader, bodyDepth) == true)
            {
                yield break;
            }

            if (skipUntil >= 0)
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == skipUntil)
                {
                    skipUntil = -1;
                }

                continue;
            }

            if (reader.NodeType == XmlNodeType.Element)
            {
                string name = reader.LocalName;
                bool empty = reader.IsEmptyElement;

                if (Skips(name, view) == true)
                {
                    if (empty == false)
                    {
                        skipUntil = reader.Depth;
                    }

                    continue;
                }

                switch (name)
                {
                    case "p":
                        yield return new DocumentEvent(DocumentEventKind.ParagraphStart, string.Empty, tables);

                        if (empty == true)
                        {
                            yield return new DocumentEvent(DocumentEventKind.ParagraphEnd, string.Empty, tables);
                        }

                        break;

                    case "tbl":
                        tables++;
                        yield return new DocumentEvent(DocumentEventKind.TableStart, string.Empty, tables - 1);

                        break;

                    case "tr":
                        yield return new DocumentEvent(DocumentEventKind.RowStart, string.Empty, tables);

                        break;

                    case "tc":
                        yield return new DocumentEvent(DocumentEventKind.CellStart, string.Empty, tables);

                        break;

                    case "fldChar":
                        OpcXml.TryGetAttribute(reader, "fldCharType", out string? type);

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

                        inInstruction = fields.Contains(true);

                        break;

                    case "t":
                    case "delText":
                        string text = await reader.ReadElementContentAsStringAsync();
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
                            OpcXml.TryGetAttribute(reader, "type", out string? breakType);

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

            if (reader.NodeType != XmlNodeType.EndElement)
            {
                continue;
            }

            switch (reader.LocalName)
            {
                case "p":
                    yield return new DocumentEvent(DocumentEventKind.ParagraphEnd, string.Empty, tables);

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
