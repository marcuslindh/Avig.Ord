using System.Globalization;
using System.Xml.Linq;

namespace Avig.Ord.Internal;

/// <summary>
/// The styles the builders use, defined as Word defines them, and the parts a new document
/// starts with. A style a builder uses is added to a document's styles part when the document
/// is saved, if the part does not have it.
/// </summary>
internal static class BuiltInStyles
{
    /// <summary>The ids of the heading styles, Heading1 … Heading9.</summary>
    public static string Heading(int level)
    {
        return "Heading" + Math.Clamp(level, 1, 9).ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>The style with this id, or false when it is not one of the built-in styles the package knows.</summary>
    public static bool TryCreate(WordNames names, string id, out XElement style)
    {
        XNamespace wml = names.W;
        style = new XElement(wml + "style");

        switch (id)
        {
            case "Normal":
                style = Paragraph(names, id, "Normal", null, true);

                return true;

            case "Title":
                style = Paragraph(names, id, "Title", "Normal", false, Run(names, size: 56, spacing: -10), ParagraphProperties(names, after: 0, contextual: true));

                return true;

            case "Subtitle":
                style = Paragraph(names, id, "Subtitle", "Normal", false, Run(names, size: 28, color: "595959"), ParagraphProperties(names, after: 160));

                return true;

            case "ListParagraph":
                style = Paragraph(names, id, "List Paragraph", "Normal", false, null, new XElement(wml + "pPr", new XElement(wml + "ind", new XAttribute(wml + "left", "720")), new XElement(wml + "contextualSpacing")));

                return true;

            case "Quote":
                style = Paragraph(names, id, "Quote", "Normal", false, Run(names, size: 22, italic: true, color: "404040"), new XElement(wml + "pPr", new XElement(wml + "spacing", new XAttribute(wml + "before", "200"), new XAttribute(wml + "after", "160")), new XElement(wml + "ind", new XAttribute(wml + "left", "864"), new XAttribute(wml + "right", "864")), new XElement(wml + "jc", new XAttribute(names.Val, "center"))));

                return true;

            case "FootnoteText":
            case "EndnoteText":
                string noteName = "footnote text";

                if (id == "EndnoteText")
                {
                    noteName = "endnote text";
                }

                style = Paragraph(names, id, noteName, "Normal", false, Run(names, size: 20), new XElement(wml + "pPr", new XElement(wml + "spacing", new XAttribute(wml + "after", "0"), new XAttribute(wml + "line", "240"), new XAttribute(wml + "lineRule", "auto"))));

                return true;

            case "DefaultParagraphFont":
                style = new XElement(
                    wml + "style",
                    new XAttribute(names.Type, "character"),
                    new XAttribute(wml + "default", "1"),
                    new XAttribute(names.StyleId, id),
                    Value(names, "name", "Default Paragraph Font"),
                    Value(names, "uiPriority", "1"),
                    new XElement(wml + "semiHidden"),
                    new XElement(wml + "unhideWhenUsed"));

                return true;

            case "Hyperlink":
                style = Character(names, id, "Hyperlink", new XElement(wml + "rPr", Value(names, "color", "0563C1"), Value(names, "u", "single")));

                return true;

            case "FootnoteReference":
            case "EndnoteReference":
                string referenceName = "footnote reference";

                if (id == "EndnoteReference")
                {
                    referenceName = "endnote reference";
                }

                style = Character(names, id, referenceName, new XElement(wml + "rPr", Value(names, "vertAlign", "superscript")));

                return true;

            case "Strong":
                style = Character(names, id, "Strong", new XElement(wml + "rPr", new XElement(wml + "b"), new XElement(wml + "bCs")));

                return true;

            case "Emphasis":
                style = Character(names, id, "Emphasis", new XElement(wml + "rPr", new XElement(wml + "i"), new XElement(wml + "iCs")));

                return true;

            case "TableNormal":
                style = new XElement(
                    wml + "style",
                    new XAttribute(names.Type, "table"),
                    new XAttribute(wml + "default", "1"),
                    new XAttribute(names.StyleId, id),
                    Value(names, "name", "Normal Table"),
                    Value(names, "uiPriority", "99"),
                    new XElement(wml + "semiHidden"),
                    new XElement(wml + "unhideWhenUsed"),
                    new XElement(
                        wml + "tblPr",
                        new XElement(wml + "tblInd", new XAttribute(wml + "w", "0"), new XAttribute(wml + "type", "dxa")),
                        new XElement(
                            wml + "tblCellMar",
                            new XElement(wml + "top", new XAttribute(wml + "w", "0"), new XAttribute(wml + "type", "dxa")),
                            new XElement(wml + "left", new XAttribute(wml + "w", "108"), new XAttribute(wml + "type", "dxa")),
                            new XElement(wml + "bottom", new XAttribute(wml + "w", "0"), new XAttribute(wml + "type", "dxa")),
                            new XElement(wml + "right", new XAttribute(wml + "w", "108"), new XAttribute(wml + "type", "dxa")))));

                return true;

            case "TableGrid":
                XElement border(string side) => new(wml + side, new XAttribute(names.Val, "single"), new XAttribute(wml + "sz", "4"), new XAttribute(wml + "space", "0"), new XAttribute(wml + "color", "auto"));

                style = new XElement(
                    wml + "style",
                    new XAttribute(names.Type, "table"),
                    new XAttribute(names.StyleId, id),
                    Value(names, "name", "Table Grid"),
                    Value(names, "basedOn", "TableNormal"),
                    Value(names, "uiPriority", "39"),
                    new XElement(wml + "pPr", new XElement(wml + "spacing", new XAttribute(wml + "after", "0"), new XAttribute(wml + "line", "240"), new XAttribute(wml + "lineRule", "auto"))),
                    new XElement(wml + "tblPr", new XElement(wml + "tblBorders", border("top"), border("left"), border("bottom"), border("right"), border("insideH"), border("insideV"))));

                return true;

            case "NoList":
                style = new XElement(
                    wml + "style",
                    new XAttribute(names.Type, "numbering"),
                    new XAttribute(wml + "default", "1"),
                    new XAttribute(names.StyleId, id),
                    Value(names, "name", "No List"),
                    Value(names, "uiPriority", "99"),
                    new XElement(wml + "semiHidden"),
                    new XElement(wml + "unhideWhenUsed"));

                return true;

            default:
                break;
        }

        if (id.Length == 8 && id.StartsWith("Heading", StringComparison.Ordinal) == true && id[7] is >= '1' and <= '9')
        {
            int level = id[7] - '0';
            int[] sizes = [32, 26, 24, 22, 22, 22, 22, 22, 22];
            int before = 240;

            if (level > 1)
            {
                before = 40;
            }

            XElement properties = new(
                wml + "pPr",
                new XElement(wml + "keepNext"),
                new XElement(wml + "keepLines"),
                new XElement(wml + "spacing", new XAttribute(wml + "before", before.ToString(CultureInfo.InvariantCulture)), new XAttribute(wml + "after", "0")),
                Value(names, "outlineLvl", (level - 1).ToString(CultureInfo.InvariantCulture)));

            style = Paragraph(names, id, "heading " + level.ToString(CultureInfo.InvariantCulture), "Normal", false, Run(names, size: sizes[level - 1], color: "2F5496", italic: level >= 4), properties, "Normal");

            return true;
        }

        return false;
    }

    /// <summary>The styles part a new document starts with: the document defaults and the styles every document has.</summary>
    public static XDocument NewStylesPart(WordNames names)
    {
        XNamespace wml = names.W;
        XElement root = new(
            wml + "styles",
            new XAttribute(XNamespace.Xmlns + "w", wml.NamespaceName),
            new XElement(
                wml + "docDefaults",
                new XElement(
                    wml + "rPrDefault",
                    new XElement(
                        wml + "rPr",
                        new XElement(wml + "rFonts", new XAttribute(wml + "ascii", "Calibri"), new XAttribute(wml + "hAnsi", "Calibri"), new XAttribute(wml + "eastAsia", "Calibri"), new XAttribute(wml + "cs", "Times New Roman")),
                        Value(names, "sz", "22"),
                        Value(names, "szCs", "22"),
                        new XElement(wml + "lang", new XAttribute(names.Val, "sv-SE"), new XAttribute(wml + "eastAsia", "en-US"), new XAttribute(wml + "bidi", "ar-SA")))),
                new XElement(
                    wml + "pPrDefault",
                    new XElement(wml + "pPr", new XElement(wml + "spacing", new XAttribute(wml + "after", "160"), new XAttribute(wml + "line", "259"), new XAttribute(wml + "lineRule", "auto"))))));

        foreach (string id in (string[])["Normal", "DefaultParagraphFont", "TableNormal", "NoList"])
        {
            if (TryCreate(names, id, out XElement style) == true)
            {
                root.Add(style);
            }
        }

        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
    }

    /// <summary>A settings part that keeps Word out of compatibility mode.</summary>
    public static XDocument NewSettingsPart(WordNames names)
    {
        XNamespace wml = names.W;

        return new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(
                wml + "settings",
                new XAttribute(XNamespace.Xmlns + "w", wml.NamespaceName),
                Value(names, "defaultTabStop", "1304"),
                new XElement(
                    wml + "compat",
                    new XElement(
                        wml + "compatSetting",
                        new XAttribute(wml + "name", "compatibilityMode"),
                        new XAttribute(wml + "uri", "http://schemas.microsoft.com/office/word"),
                        new XAttribute(wml + "val", "15")))));
    }

    /// <summary>A footnotes or endnotes part with the two separators Word expects first in it.</summary>
    public static XDocument NewNotesPart(WordNames names, NoteKind kind)
    {
        XNamespace wml = names.W;
        string element = "footnote";
        string root = "footnotes";

        if (kind == NoteKind.Endnote)
        {
            element = "endnote";
            root = "endnotes";
        }

        XElement separator(string type, string id, string mark) => new(
            wml + element,
            new XAttribute(names.Type, type),
            new XAttribute(names.WId, id),
            new XElement(
                wml + "p",
                new XElement(wml + "pPr", new XElement(wml + "spacing", new XAttribute(wml + "after", "0"), new XAttribute(wml + "line", "240"), new XAttribute(wml + "lineRule", "auto"))),
                new XElement(wml + "r", new XElement(wml + mark))));

        return new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(
                wml + root,
                new XAttribute(XNamespace.Xmlns + "w", wml.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "r", names.R.NamespaceName),
                separator("separator", "-1", "separator"),
                separator("continuationSeparator", "0", "continuationSeparator")));
    }

    /// <summary>
    /// A numbering part with the two list definitions the builders use: abstract 0 bulleted
    /// (•, ◦, ▪ by level) and abstract 1 numbered (1., a., i., then again).
    /// </summary>
    public static XDocument NewNumberingPart(WordNames names)
    {
        XNamespace wml = names.W;
        string[] bullets = ["•", "◦", "▪"];
        (string Format, string Suffix)[] numbers = [("decimal", "."), ("lowerLetter", "."), ("lowerRoman", ".")];
        XElement bulleted = new(wml + "abstractNum", new XAttribute(wml + "abstractNumId", "0"), Value(names, "multiLevelType", "hybridMultilevel"));
        XElement numbered = new(wml + "abstractNum", new XAttribute(wml + "abstractNumId", "1"), Value(names, "multiLevelType", "hybridMultilevel"));

        for (int level = 0; level < 9; level++)
        {
            string indent = (720 * (level + 1)).ToString(CultureInfo.InvariantCulture);
            string levelText = level.ToString(CultureInfo.InvariantCulture);
            XElement properties() => new(wml + "pPr", new XElement(wml + "ind", new XAttribute(wml + "left", indent), new XAttribute(wml + "hanging", "360")));

            bulleted.Add(new XElement(
                wml + "lvl",
                new XAttribute(wml + "ilvl", levelText),
                Value(names, "start", "1"),
                Value(names, "numFmt", "bullet"),
                Value(names, "lvlText", bullets[level % 3]),
                Value(names, "lvlJc", "left"),
                properties()));

            (string format, string suffix) = numbers[level % 3];
            numbered.Add(new XElement(
                wml + "lvl",
                new XAttribute(wml + "ilvl", levelText),
                Value(names, "start", "1"),
                Value(names, "numFmt", format),
                Value(names, "lvlText", "%" + (level + 1).ToString(CultureInfo.InvariantCulture) + suffix),
                Value(names, "lvlJc", "left"),
                properties()));
        }

        return new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(wml + "numbering", new XAttribute(XNamespace.Xmlns + "w", wml.NamespaceName), bulleted, numbered));
    }

    /// <summary>A w:num for one of the two definitions; restarting numbered lists at 1.</summary>
    public static XElement NumberingInstance(WordNames names, string numberingId, bool bullet)
    {
        XNamespace wml = names.W;
        string abstractId = "1";

        if (bullet == true)
        {
            abstractId = "0";
        }

        XElement instance = new(wml + "num", new XAttribute(wml + "numId", numberingId), Value(names, "abstractNumId", abstractId));

        if (bullet == false)
        {
            instance.Add(new XElement(wml + "lvlOverride", new XAttribute(wml + "ilvl", "0"), Value(names, "startOverride", "1")));
        }

        return instance;
    }

    private static XElement Value(WordNames names, string name, string value)
    {
        return new XElement(names.W + name, new XAttribute(names.Val, value));
    }

    private static XElement Paragraph(WordNames names, string id, string name, string? basedOn, bool isDefault, XElement? run = null, XElement? paragraph = null, string? next = null)
    {
        XNamespace wml = names.W;
        XElement style = new(wml + "style", new XAttribute(names.Type, "paragraph"));

        if (isDefault == true)
        {
            style.Add(new XAttribute(wml + "default", "1"));
        }

        style.Add(new XAttribute(names.StyleId, id), Value(names, "name", name));

        if (basedOn is not null)
        {
            style.Add(Value(names, "basedOn", basedOn));
        }

        if (next is not null)
        {
            style.Add(Value(names, "next", next));
        }

        style.Add(new XElement(wml + "qFormat"));

        if (paragraph is not null)
        {
            style.Add(paragraph);
        }

        if (run is not null)
        {
            style.Add(run);
        }

        return style;
    }

    private static XElement Character(WordNames names, string id, string name, XElement run)
    {
        XNamespace wml = names.W;

        return new XElement(
            wml + "style",
            new XAttribute(names.Type, "character"),
            new XAttribute(names.StyleId, id),
            Value(names, "name", name),
            Value(names, "basedOn", "DefaultParagraphFont"),
            Value(names, "uiPriority", "99"),
            new XElement(wml + "unhideWhenUsed"),
            run);
    }

    private static XElement Run(WordNames names, int size, string? color = null, bool italic = false, int? spacing = null)
    {
        XNamespace wml = names.W;
        XElement run = new(wml + "rPr");

        if (italic == true)
        {
            run.Add(new XElement(wml + "i"), new XElement(wml + "iCs"));
        }

        if (color is not null)
        {
            run.Add(Value(names, "color", color));
        }

        if (spacing is not null)
        {
            run.Add(Value(names, "spacing", spacing.Value.ToString(CultureInfo.InvariantCulture)));
        }

        string halfPoints = size.ToString(CultureInfo.InvariantCulture);
        run.Add(Value(names, "sz", halfPoints), Value(names, "szCs", halfPoints));

        return run;
    }

    private static XElement ParagraphProperties(WordNames names, int after, bool contextual = false)
    {
        XNamespace wml = names.W;
        XElement properties = new(wml + "pPr");
        properties.Add(new XElement(wml + "spacing", new XAttribute(wml + "after", after.ToString(CultureInfo.InvariantCulture)), new XAttribute(wml + "line", "240"), new XAttribute(wml + "lineRule", "auto")));

        if (contextual == true)
        {
            properties.Add(new XElement(wml + "contextualSpacing"));
        }

        return properties;
    }
}
