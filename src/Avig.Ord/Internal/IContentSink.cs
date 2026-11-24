using System.Xml.Linq;

namespace Avig.Ord.Internal;

/// <summary>
/// What the builders need from whoever puts their content in a file: relationships for links
/// and pictures, notes, styles and lists. <see cref="DocxWriter"/> collects them for the parts
/// it writes at the end; <see cref="OrdDocument"/> adds them to its parts straight away.
/// </summary>
internal interface IContentSink
{
    WordNames Names { get; }

    /// <summary>The width between the margins, for tables and pictures that should fill it.</summary>
    Length TextWidth { get; }

    /// <summary>A relationship id for a hyperlink to a URL.</summary>
    string AddHyperlink(string url);

    /// <summary>A relationship id for a picture whose bytes go into the package.</summary>
    string AddPicture(byte[] data, string contentType, string extension);

    /// <summary>A number for a drawing's wp:docPr/@id, unique in the document.</summary>
    int NextDrawingId();

    /// <summary>Adds a footnote or endnote with these paragraphs and gives back its id.</summary>
    string AddNote(NoteKind kind, IReadOnlyList<XElement> paragraphs);

    /// <summary>Says that a style is used, so that it is in the styles part when the file is saved.</summary>
    void UseStyle(string styleId);

    /// <summary>The w:numId of the bulleted or numbered list a paragraph joins; a new list when <paramref name="startNew"/>.</summary>
    string ListNumbering(bool bullet, bool startNew);
}
