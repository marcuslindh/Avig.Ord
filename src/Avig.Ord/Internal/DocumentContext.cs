using System.Runtime.CompilerServices;
using System.Xml.Linq;
using Avig.OpenXml;

namespace Avig.Ord.Internal;

/// <summary>
/// What a paragraph or table needs to know about the document it is in: the names of its
/// namespace, the styles, the list labels, the relationships of the part it is in (for
/// hyperlinks and pictures), how tracked changes are read, and whether it may be changed.
/// </summary>
internal sealed class DocumentContext
{
    private readonly ConditionalWeakTable<XElement, ListInfo> listLabels = [];
    private readonly Func<string, CancellationToken, Task<Result<byte[]>>> readPart;
    private readonly StyleSheet styles;
    private readonly NumberingDefinitions numbering;
    private ListCounter counter;

    public DocumentContext(
        WordNames names,
        StyleSheet styles,
        NumberingDefinitions numbering,
        string partName,
        IReadOnlyList<OpcRelationship> relationships,
        RevisionView view,
        OrdDocument? owner,
        Func<string, CancellationToken, Task<Result<byte[]>>> readPart)
    {
        Names = names;
        this.styles = styles;
        this.numbering = numbering;
        PartName = partName;
        Relationships = relationships;
        View = view;
        Owner = owner;
        this.readPart = readPart;
        counter = new ListCounter(numbering);
    }

    public WordNames Names { get; }

    /// <summary>The styles: the document's current ones (they change when the package adds a style), or those read with the streaming reader.</summary>
    public StyleSheet Styles => Owner?.Styles ?? styles;

    public NumberingDefinitions Numbering => Owner?.Numbering ?? numbering;

    /// <summary>The part the content is in: word/document.xml, a header, the footnotes …</summary>
    public string PartName { get; }

    public IReadOnlyList<OpcRelationship> Relationships { get; set; }

    public RevisionView View { get; }

    /// <summary>The document model the content belongs to; null when it was read by the streaming reader and cannot be changed.</summary>
    public OrdDocument? Owner { get; }

    public bool IsReadOnly => Owner is null;

    /// <summary>Changed every time the content is changed, so that list labels counted before are counted again.</summary>
    public int Version { get; private set; }

    public int LabelledVersion { get; set; } = -1;

    public void ThrowIfReadOnly()
    {
        if (IsReadOnly == true)
        {
            throw new InvalidOperationException("This content was read with DocxReader, which reads a document without loading it, and cannot be changed. Open the document with OrdDocument.OpenAsync to change it.");
        }
    }

    public void Changed()
    {
        Version++;
    }

    public Task<Result<byte[]>> ReadPartAsync(string partName, CancellationToken cancellationToken)
    {
        return readPart(partName, cancellationToken);
    }

    /// <summary>The relationship with this id, resolved: a part name for an internal one, a URL for an external one.</summary>
    public bool TryResolve(string id, out string target, out bool external)
    {
        target = string.Empty;
        external = false;

        foreach (OpcRelationship relationship in Relationships)
        {
            if (string.Equals(relationship.Id, id, StringComparison.Ordinal) == false)
            {
                continue;
            }

            external = relationship.IsExternal;
            target = relationship.Target;

            if (external == false)
            {
                target = OpcPath.ResolveTarget(PartName, relationship.Target);
            }

            return true;
        }

        return false;
    }

    /// <summary>Counts the paragraph as the next list item, in document order (the streaming reader calls this as it goes).</summary>
    public void CountListItem(XElement paragraph, string? styleId, XElement? properties)
    {
        (string? numberingId, int? level) = NumberingOf(styleId, properties);

        if (numberingId is not null && counter.TryNext(numberingId, level ?? 0, out ListInfo info) == true)
        {
            listLabels.AddOrUpdate(paragraph, info);
        }
    }

    public bool TryGetListLabel(XElement paragraph, out ListInfo? info)
    {
        return listLabels.TryGetValue(paragraph, out info);
    }

    /// <summary>Forgets the list labels counted so far, before they are counted again from the start of the document.</summary>
    public void ResetListCounts()
    {
        listLabels.Clear();
        counter = new ListCounter(Numbering);
    }

    private (string? NumberingId, int? Level) NumberingOf(string? styleId, XElement? properties)
    {
        XElement? direct = properties?.Element(Names.W + "numPr");
        (string? styleNumbering, int? styleLevel) = Styles.StyleNumbering(styleId);
        string? numberingId = Names.ValueOf(direct, "numId") ?? styleNumbering;
        int? level = FormatReader.Integer(Names.ValueOf(direct, "ilvl")) ?? styleLevel;

        return (numberingId, level);
    }
}
