using System.Xml.Linq;

namespace Avig.Ord.Internal;

/// <summary>
/// The element and attribute names of WordprocessingML in the namespace a document uses.
/// Transitional and Strict OOXML have the same local names in different namespaces, so a
/// document's names are taken from its root element, and every lookup goes through them. That
/// way a Strict document is read, edited and saved as Strict.
/// </summary>
internal sealed class WordNames
{
    public const string TransitionalMain = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    public const string StrictMain = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    public const string TransitionalRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public const string StrictRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    public const string Drawing = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    public const string DrawingMain = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public const string Picture = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    public const string SvgBlip = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
    public const string Word2010 = "http://schemas.microsoft.com/office/word/2010/wordml";
    public const string Vml = "urn:schemas-microsoft-com:vml";
    public const string Xml = "http://www.w3.org/XML/1998/namespace";

    private WordNames(XNamespace main, XNamespace relationships)
    {
        W = main;
        R = relationships;
        Body = W + "body";
        Paragraph = W + "p";
        ParagraphProperties = W + "pPr";
        Run = W + "r";
        RunProperties = W + "rPr";
        Text = W + "t";
        DeletedText = W + "delText";
        InstructionText = W + "instrText";
        DeletedInstructionText = W + "delInstrText";
        Table = W + "tbl";
        TableRow = W + "tr";
        TableCell = W + "tc";
        TableProperties = W + "tblPr";
        TableGrid = W + "tblGrid";
        TableRowProperties = W + "trPr";
        TableCellProperties = W + "tcPr";
        Hyperlink = W + "hyperlink";
        Inserted = W + "ins";
        Deleted = W + "del";
        MoveFrom = W + "moveFrom";
        MoveTo = W + "moveTo";
        ContentControl = W + "sdt";
        ContentControlContent = W + "sdtContent";
        ContentControlProperties = W + "sdtPr";
        SimpleField = W + "fldSimple";
        FieldChar = W + "fldChar";
        SectionProperties = W + "sectPr";
        Val = W + "val";
        Id = R + "id";
        Embed = R + "embed";
        Link = R + "link";
        WId = W + "id";
        Type = W + "type";
        Style = W + "style";
        StyleId = W + "styleId";
    }

    public static WordNames Transitional { get; } = new(TransitionalMain, TransitionalRelationships);

    public static WordNames Strict { get; } = new(StrictMain, StrictRelationships);

    public XNamespace W { get; }

    public XNamespace R { get; }

    public bool IsStrict => W == StrictMain;

    public XName Body { get; }

    public XName Paragraph { get; }

    public XName ParagraphProperties { get; }

    public XName Run { get; }

    public XName RunProperties { get; }

    public XName Text { get; }

    public XName DeletedText { get; }

    public XName InstructionText { get; }

    public XName DeletedInstructionText { get; }

    public XName Table { get; }

    public XName TableRow { get; }

    public XName TableCell { get; }

    public XName TableProperties { get; }

    public XName TableGrid { get; }

    public XName TableRowProperties { get; }

    public XName TableCellProperties { get; }

    public XName Hyperlink { get; }

    public XName Inserted { get; }

    public XName Deleted { get; }

    public XName MoveFrom { get; }

    public XName MoveTo { get; }

    public XName ContentControl { get; }

    public XName ContentControlContent { get; }

    public XName ContentControlProperties { get; }

    public XName SimpleField { get; }

    public XName FieldChar { get; }

    public XName SectionProperties { get; }

    /// <summary>w:val, the attribute most properties keep their value in.</summary>
    public XName Val { get; }

    /// <summary>r:id.</summary>
    public XName Id { get; }

    /// <summary>r:embed, on a picture's blip.</summary>
    public XName Embed { get; }

    /// <summary>r:link, on a picture that is linked rather than embedded.</summary>
    public XName Link { get; }

    /// <summary>w:id, on notes, comments and revisions.</summary>
    public XName WId { get; }

    /// <summary>w:type, on breaks, notes and styles.</summary>
    public XName Type { get; }

    public XName Style { get; }

    public XName StyleId { get; }

    /// <summary>The names for a part whose root element is <paramref name="root"/>: Strict when its namespace is Strict's.</summary>
    public static WordNames For(XElement? root)
    {
        if (root is not null && root.Name.Namespace == StrictMain)
        {
            return Strict;
        }

        return Transitional;
    }

    /// <summary>A name in the main namespace.</summary>
    public XName Name(string localName)
    {
        return W + localName;
    }

    /// <summary>The w:val of a child element, such as w:pStyle/@w:val.</summary>
    public string? ValueOf(XElement? parent, string childName)
    {
        return parent?.Element(W + childName)?.Attribute(Val)?.Value;
    }

    /// <summary>
    /// The on/off state of a toggle property such as w:b: present without w:val, or with true,
    /// 1 or on, is on; false, 0, off and none are off; null when the element is not there.
    /// </summary>
    public bool? ToggleOf(XElement? properties, string childName)
    {
        XElement? element = properties?.Element(W + childName);
        bool? state = null;

        if (element is not null)
        {
            string? value = element.Attribute(Val)?.Value;
            state = value is null || value is not ("false" or "0" or "off" or "none");
        }

        return state;
    }
}
