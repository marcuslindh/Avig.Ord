using System.Globalization;
using System.Security;
using System.Text;

namespace Avig.Ord.Fuzz;

/// <summary>
/// Makes the content of a w:body from a grammar of what Word writes: paragraphs with styles and
/// list numbering, runs with formatting, tabs, breaks and symbols, hyperlinks, simple and
/// complex fields (nested, and now and then without their end or begin), tracked insertions,
/// deletions and moves, content controls, smart tags, text boxes, pictures, note references,
/// and tables with merged cells and tables inside them. Random bytes never become a body that
/// gets past the first element, so the targets that need structure generate it.
/// </summary>
internal sealed class BodyGenerator
{
    private static readonly IReadOnlyDictionary<string, string> Ids = Template.MainRelationshipIds();

    private static readonly string[] ParagraphStyles = ["Heading1", "Heading2", "Heading3", "Normal", "ListParagraph", "Title", "Unknown", string.Empty];

    private static readonly string[] Words = ["Det", "var", "en", "ny", "text", "Öst", "1.", "*fet*", "[a]", "#", "|", "_", "a  b", " ", "\t", "&<>\"'", "😀", "ﬁ"];

    private readonly Random random;
    private readonly StringBuilder xml = new();
    private int id = 100;

    private BodyGenerator(Random random)
    {
        this.random = random;
    }

    /// <summary>A body of a few blocks.</summary>
    public static string Body(Random random)
    {
        BodyGenerator generator = new(random);
        int blocks = random.Next(1, 8);

        for (int index = 0; index < blocks; index++)
        {
            generator.Block(0);
        }

        return generator.xml.ToString();
    }

    private void Block(int depth)
    {
        int kind = random.Next(20);

        if (kind < 12 || depth > 3)
        {
            Paragraph(depth);
        }
        else if (kind < 15)
        {
            Table(depth);
        }
        else if (kind < 17)
        {
            xml.Append("<w:sdt><w:sdtPr><w:tag w:val=\"t").Append(id++).Append("\"/></w:sdtPr><w:sdtContent>");
            Block(depth + 1);
            xml.Append("</w:sdtContent></w:sdt>");
        }
        else if (kind == 17)
        {
            xml.Append("<w:customXml w:element=\"x\">");
            Block(depth + 1);
            xml.Append("</w:customXml>");
        }
        else
        {
            string revision = Pick(["ins", "del"]);
            xml.Append("<w:").Append(revision).Append(Revision()).Append('>');
            Paragraph(depth, revision == "del");
            xml.Append("</w:").Append(revision).Append('>');
        }
    }

    private void Paragraph(int depth, bool deleted = false)
    {
        xml.Append("<w:p>");
        ParagraphProperties();
        int inlines = random.Next(0, 6);

        for (int index = 0; index < inlines; index++)
        {
            Inline(depth, deleted, 0);
        }

        xml.Append("</w:p>");
    }

    private void ParagraphProperties()
    {
        if (random.Next(3) > 0)
        {
            return;
        }

        xml.Append("<w:pPr>");

        if (random.Next(2) == 0)
        {
            xml.Append("<w:pStyle w:val=\"").Append(Pick(ParagraphStyles)).Append("\"/>");
        }

        if (random.Next(2) == 0)
        {
            xml.Append("<w:numPr><w:ilvl w:val=\"").Append(Pick(["0", "1", "2", "8", "9", "-1"])).Append("\"/><w:numId w:val=\"").Append(Pick(["1", "2", "3", "0", "99"])).Append("\"/></w:numPr>");
        }

        if (random.Next(6) == 0)
        {
            xml.Append("<w:outlineLvl w:val=\"").Append(random.Next(-1, 11)).Append("\"/>");
        }

        if (random.Next(8) == 0)
        {
            xml.Append("<w:rPr><w:ins").Append(Revision()).Append("/></w:rPr>");
        }

        xml.Append("</w:pPr>");
    }

    private void Inline(int depth, bool deleted, int nesting)
    {
        int kind = random.Next(30);

        if (nesting > 3)
        {
            kind = 0;
        }

        switch (kind)
        {
            case < 12:
                Run(deleted);

                break;

            case < 14:
                xml.Append("<w:hyperlink");

                if (random.Next(2) == 0)
                {
                    xml.Append(" r:id=\"").Append(Pick([Ids["hyperlink"], "rId999", string.Empty])).Append('"');
                }
                else
                {
                    xml.Append(" w:anchor=\"_Toc").Append(id++).Append('"');
                }

                xml.Append('>');
                Inline(depth, deleted, nesting + 1);
                xml.Append("</w:hyperlink>");

                break;

            case < 16:
                xml.Append("<w:fldSimple w:instr=\"").Append(Escape(Pick([" PAGE ", "MERGEFIELD Namn", "HYPERLINK \"https://example.com/\"", " INCLUDEPICTURE \"http://example.com/a.png\" \\d", string.Empty]))).Append("\">");
                Run(deleted);
                xml.Append("</w:fldSimple>");

                break;

            case < 20:
                ComplexField(depth, deleted, nesting);

                break;

            case < 22:
                xml.Append("<w:sdt><w:sdtPr><w:alias w:val=\"Namn\"/><w:tag w:val=\"namn\"/></w:sdtPr><w:sdtContent>");
                Inline(depth, deleted, nesting + 1);
                xml.Append("</w:sdtContent></w:sdt>");

                break;

            case < 25:
                string revision = Pick(["ins", "del", "moveFrom", "moveTo"]);
                xml.Append("<w:").Append(revision).Append(Revision()).Append('>');
                Run(deleted || revision is "del" or "moveFrom");
                xml.Append("</w:").Append(revision).Append('>');

                break;

            case 25:
                xml.Append("<w:smartTag w:uri=\"urn:x\" w:element=\"x\">");
                Inline(depth, deleted, nesting + 1);
                xml.Append("</w:smartTag>");

                break;

            case 26:
                TextBox(depth);

                break;

            case 27:
                xml.Append("<w:r><w:drawing><wp:inline><wp:extent cx=\"").Append(Pick(["914400", "-1", "0", "9223372036854775807"])).Append("\" cy=\"457200\"/><wp:docPr id=\"").Append(id++)
                    .Append("\" name=\"Bild\" descr=\"").Append(Escape(Pick(Words))).Append("\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:blipFill><a:blip r:embed=\"")
                    .Append(Pick([Ids["image"], "rId999", Ids["styles"]])).Append("\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>");

                break;

            case 28:
                xml.Append("<w:r><w:footnoteReference w:id=\"").Append(Pick(["1", "2", "99"])).Append("\"/></w:r>");

                break;

            default:
                xml.Append("<w:bookmarkStart w:id=\"").Append(id).Append("\" w:name=\"b").Append(id).Append("\"/>");
                Run(deleted);
                xml.Append("<w:bookmarkEnd w:id=\"").Append(id++).Append("\"/>");

                break;
        }
    }

    /// <summary>A complex field: begin, instruction, separate, result, end. Now and then a part is missing or a field nests in the instruction or the result.</summary>
    private void ComplexField(int depth, bool deleted, int nesting)
    {
        bool begin = random.Next(12) > 0;
        bool separate = random.Next(6) > 0;
        bool end = random.Next(12) > 0;

        if (begin == true)
        {
            xml.Append("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>");
        }

        string instruction = Pick(["MERGEFIELD Förnamn \\* MERGEFORMAT", "PAGE", "HYPERLINK \"https://example.com/\"", "IF ", "DATE", "TOC \\o \"1-3\"", "DDEAUTO c:\\\\x.exe \"/k\"", "INCLUDETEXT \"a.docx\""]);
        xml.Append("<w:r><w:").Append(Instruction(deleted)).Append(" xml:space=\"preserve\">").Append(Escape(instruction)).Append("</w:").Append(Instruction(deleted)).Append("></w:r>");

        if (random.Next(5) == 0)
        {
            ComplexField(depth, deleted, nesting + 1);
        }

        if (separate == true)
        {
            xml.Append("<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>");
        }

        Inline(depth, deleted, nesting + 1);

        if (end == true)
        {
            xml.Append("<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>");
        }
    }

    private static string Instruction(bool deleted)
    {
        if (deleted == true)
        {
            return "delInstrText";
        }

        return "instrText";
    }

    private void TextBox(int depth)
    {
        xml.Append("<w:r><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" Requires=\"wps\"><w:drawing><wps:txbx><w:txbxContent>");
        Block(depth + 2);
        xml.Append("</w:txbxContent></wps:txbx></w:drawing></mc:Choice><mc:Fallback><w:pict><w:p><w:r><w:t>Reserv</w:t></w:r></w:p></w:pict></mc:Fallback></mc:AlternateContent></w:r>");
    }

    private void Run(bool deleted)
    {
        xml.Append("<w:r>");

        if (random.Next(3) == 0)
        {
            xml.Append("<w:rPr>");
            Toggle("b");
            Toggle("i");
            Toggle("strike");
            Toggle("caps");

            if (random.Next(4) == 0)
            {
                xml.Append("<w:rStyle w:val=\"").Append(Pick(["Hyperlink", "Strong", "Unknown"])).Append("\"/>");
            }

            if (random.Next(4) == 0)
            {
                xml.Append("<w:sz w:val=\"").Append(Pick(["24", "0", "-2", "3277", "99999999999"])).Append("\"/>");
            }

            xml.Append("</w:rPr>");
        }

        int pieces = random.Next(1, 4);

        for (int piece = 0; piece < pieces; piece++)
        {
            int kind = random.Next(14);

            if (kind < 8)
            {
                string text = Escape(Pick(Words));

                if (deleted == true)
                {
                    xml.Append("<w:delText xml:space=\"preserve\">").Append(text).Append("</w:delText>");
                }
                else
                {
                    xml.Append("<w:t xml:space=\"preserve\">").Append(text).Append("</w:t>");
                }
            }
            else if (kind == 8)
            {
                xml.Append("<w:tab/>");
            }
            else if (kind == 9)
            {
                xml.Append("<w:br").Append(Pick([string.Empty, " w:type=\"page\"", " w:type=\"column\"", " w:type=\"textWrapping\""])).Append("/>");
            }
            else if (kind == 10)
            {
                xml.Append("<w:sym w:font=\"Wingdings\" w:char=\"").Append(Pick(["F04A", "0041", "FFFFFFFF", "D800", "110000", "0", "zz"])).Append("\"/>");
            }
            else if (kind == 11)
            {
                xml.Append("<w:noBreakHyphen/>");
            }
            else if (kind == 12)
            {
                xml.Append("<w:cr/>");
            }
            else
            {
                xml.Append("<w:softHyphen/>");
            }
        }

        xml.Append("</w:r>");
    }

    private void Toggle(string name)
    {
        if (random.Next(3) == 0)
        {
            xml.Append("<w:").Append(name).Append(Pick([string.Empty, " w:val=\"0\"", " w:val=\"true\"", " w:val=\"off\""])).Append("/>");
        }
    }

    private void Table(int depth)
    {
        int rows = random.Next(1, 4);
        int columns = random.Next(1, 4);
        xml.Append("<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/></w:tblPr>");

        if (random.Next(4) > 0)
        {
            xml.Append("<w:tblGrid>");

            for (int column = 0; column < columns; column++)
            {
                xml.Append("<w:gridCol w:w=\"2000\"/>");
            }

            xml.Append("</w:tblGrid>");
        }

        for (int row = 0; row < rows; row++)
        {
            xml.Append("<w:tr>");

            if (random.Next(5) == 0)
            {
                xml.Append("<w:trPr>").Append(Pick(["<w:tblHeader/>", "<w:ins" + Revision() + "/>", "<w:del" + Revision() + "/>"])).Append("</w:trPr>");
            }

            for (int column = 0; column < columns; column++)
            {
                xml.Append("<w:tc>");

                if (random.Next(4) == 0)
                {
                    xml.Append("<w:tcPr>").Append(Pick(["<w:gridSpan w:val=\"2\"/>", "<w:vMerge w:val=\"restart\"/>", "<w:vMerge/>", "<w:gridSpan w:val=\"0\"/>"])).Append("</w:tcPr>");
                }

                Block(depth + 1);

                if (random.Next(3) == 0)
                {
                    Paragraph(depth + 1);
                }

                xml.Append("</w:tc>");
            }

            xml.Append("</w:tr>");
        }

        xml.Append("</w:tbl>");
    }

    private string Revision()
    {
        return string.Create(CultureInfo.InvariantCulture, $" w:id=\"{id++}\" w:author=\"A\" w:date=\"2026-11-24T00:00:00Z\"");
    }

    private string Pick(string[] items)
    {
        return items[random.Next(items.Length)];
    }

    private static string Escape(string text)
    {
        return SecurityElement.Escape(text);
    }
}
