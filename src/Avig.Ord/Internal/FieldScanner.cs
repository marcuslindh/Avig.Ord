using System.Text;
using System.Xml;
using System.Xml.Linq;
using Avig.OpenXml;

namespace Avig.Ord.Internal;

/// <summary>
/// Finds fields by their instruction across a whole part, as Word does: a complex field runs
/// from a fldChar begin to its end, its instruction is the instrText up to its separate, and
/// fields nest. Used to find and to neutralize fields that start programs or fetch things.
/// </summary>
internal static class FieldScanner
{
    /// <summary>A complex field: its three markers, its instruction, and every marker and instruction element inside its instruction (those of nested fields too).</summary>
    internal sealed class ComplexField(XElement begin)
    {
        public XElement Begin { get; } = begin;

        public XElement? Separate { get; set; }

        public XElement? End { get; set; }

        public StringBuilder Instruction { get; } = new();

        public List<XElement> InstructionNodes { get; } = [];
    }

    /// <summary>The complex fields of a part that are closed, in the order they end.</summary>
    public static List<ComplexField> ComplexFields(WordNames names, XElement root)
    {
        List<ComplexField> found = [];
        List<ComplexField> open = [];
        XName fieldChar = names.W + "fldChar";
        XName type = names.W + "fldCharType";

        foreach (XElement element in root.Descendants())
        {
            bool instruction = element.Name == names.InstructionText || element.Name == names.DeletedInstructionText;

            if (element.Name != fieldChar && instruction == false)
            {
                continue;
            }

            string? kind = (string?)element.Attribute(type);

            if (element.Name == fieldChar && kind == "begin")
            {
                MarkInstruction(open, element);

                // Fields nested past the cap are part of the one they are in, as the readers have it.
                if (open.Count < InlineReader.MaxFieldNesting)
                {
                    open.Add(new ComplexField(element));
                }

                continue;
            }

            if (open.Count == 0)
            {
                continue;
            }

            ComplexField current = open[^1];

            if (instruction == true)
            {
                if (current.Separate is null)
                {
                    current.Instruction.Append(element.Value);
                }

                MarkInstruction(open, element);
            }
            else if (kind == "separate" && current.Separate is null)
            {
                MarkInstruction(open.Take(open.Count - 1), element);
                current.Separate = element;
            }
            else if (kind == "end")
            {
                open.RemoveAt(open.Count - 1);
                MarkInstruction(open, element);
                current.End = element;
                found.Add(current);
            }
        }

        return found;
    }

    /// <summary>True when the part has a field that starts a program or fetches something.</summary>
    public static bool HasActiveField(WordNames names, XElement root)
    {
        XName instr = names.W + "instr";

        if (root.Descendants(names.W + "fldSimple").Any(field => ActiveContentRules.IsActiveField((string?)field.Attribute(instr) ?? string.Empty) == true) == true)
        {
            return true;
        }

        return ComplexFields(names, root).Any(field => ActiveContentRules.IsActiveField(field.Instruction.ToString()) == true);
    }

    /// <summary>
    /// Turns every field that starts a program or fetches something into the text Word last
    /// showed for it: a simple field gives way to its content, a complex one loses its markers
    /// and instruction and keeps its result.
    /// </summary>
    public static bool Neutralize(WordNames names, XElement root)
    {
        bool changed = false;
        XName instr = names.W + "instr";

        foreach (XElement simple in root.Descendants(names.W + "fldSimple").ToList())
        {
            if (simple.Parent is not null && ActiveContentRules.IsActiveField((string?)simple.Attribute(instr) ?? string.Empty) == true)
            {
                simple.ReplaceWith(simple.Nodes());
                changed = true;
            }
        }

        foreach (ComplexField field in ComplexFields(names, root))
        {
            if (ActiveContentRules.IsActiveField(field.Instruction.ToString()) == false)
            {
                continue;
            }

            foreach (XElement node in field.InstructionNodes.Append(field.Begin).Append(field.Separate).Append(field.End).OfType<XElement>())
            {
                if (node.Parent is not null)
                {
                    node.Remove();
                }
            }

            changed = true;
        }

        return changed;
    }

    /// <summary>The same test as <see cref="HasActiveField(WordNames, XElement)"/>, on a part read as a stream.</summary>
    public static async Task<bool> HasActiveFieldAsync(XmlReader reader, CancellationToken cancellationToken)
    {
        List<(StringBuilder Instruction, bool Separated)> open = [];
        int instructionDepth = -1;

        while (await reader.ReadAsync() == true)
        {
            cancellationToken.ThrowIfCancellationRequested();

            // The text of an instruction, read node by node so that nothing after it is skipped.
            if (instructionDepth >= 0)
            {
                if (reader.NodeType is XmlNodeType.Text or XmlNodeType.SignificantWhitespace or XmlNodeType.Whitespace or XmlNodeType.CDATA && open.Count > 0 && open[^1].Separated == false)
                {
                    open[^1].Instruction.Append(reader.Value);
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == instructionDepth)
                {
                    instructionDepth = -1;
                }

                continue;
            }

            if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI is not (WordNames.TransitionalMain or WordNames.StrictMain))
            {
                continue;
            }

            switch (reader.LocalName)
            {
                case "fldSimple":
                    if (OpcXml.TryGetAttribute(reader, "instr", out string? simple) == true && ActiveContentRules.IsActiveField(simple) == true)
                    {
                        return true;
                    }

                    break;

                case "fldChar":
                    OpcXml.TryGetAttribute(reader, "fldCharType", out string? kind);

                    if (kind == "begin")
                    {
                        if (open.Count < InlineReader.MaxFieldNesting)
                        {
                            open.Add((new StringBuilder(), false));
                        }
                    }
                    else if (kind == "separate" && open.Count > 0)
                    {
                        open[^1] = (open[^1].Instruction, true);
                    }
                    else if (kind == "end" && open.Count > 0)
                    {
                        if (ActiveContentRules.IsActiveField(open[^1].Instruction.ToString()) == true)
                        {
                            return true;
                        }

                        open.RemoveAt(open.Count - 1);
                    }

                    break;

                case "instrText":
                case "delInstrText":
                    if (reader.IsEmptyElement == false)
                    {
                        instructionDepth = reader.Depth;
                    }

                    break;

                default:
                    break;
            }
        }

        return false;
    }

    private static void MarkInstruction(IEnumerable<ComplexField> fields, XElement element)
    {
        foreach (ComplexField field in fields)
        {
            if (field.Separate is null)
            {
                field.InstructionNodes.Add(element);
            }
        }
    }
}
