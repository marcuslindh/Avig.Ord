using System.Xml.Linq;
using Avig.OpenXml;
using Avig.Ord.Internal;

namespace Avig.Ord;

public sealed partial class OrdDocument
{
    /// <summary>The relationship namespaces of transitional and Strict OOXML, where r:id and its kin live.</summary>
    private static readonly string[] RelationshipNamespaces = [OpcNamespaces.Relationships, "http://purl.oclc.org/ooxml/officeDocument/relationships"];

    /// <summary>
    /// Elements that are nothing but a reference to what <see cref="RemoveActiveContent"/> takes
    /// out, and go with it: an ActiveX control, an OLE object, an imported chunk, an attached
    /// template, a subdocument, a frame. Other elements only lose the attribute (a linked picture
    /// keeps its embedded one, if it has it).
    /// </summary>
    private static readonly string[] ReferenceOnlyElements = ["control", "OLEObject", "altChunk", "attachedTemplate", "subDoc", "frame"];

    /// <summary>
    /// What in the document can act on its own or fetch something when Word opens it: macros,
    /// ActiveX controls, embedded objects, fields such as DDE, and external references other
    /// than hyperlinks. None of it is ever run or followed here, but it is saved again unless
    /// <see cref="RemoveActiveContent"/> takes it out.
    /// </summary>
    public ActiveContent ActiveContent
    {
        get
        {
            ActiveContent found = ActiveContent.None;

            foreach ((string name, PartData part) in parts)
            {
                found |= ActiveContentRules.OfPart(name, part.ContentType);
            }

            foreach ((string source, List<OpcRelationship> list) in relationships)
            {
                if (source.Length > 0 && parts.ContainsKey(source) == false)
                {
                    continue;
                }

                foreach (OpcRelationship relationship in list)
                {
                    found |= ActiveContentRules.OfRelationship(relationship);
                }
            }

            if (HasActiveFields() == true)
            {
                found |= ActiveContent.ActiveFields;
            }

            return found;
        }
    }

    /// <summary>
    /// The relationships to something outside the file that Word follows on its own, with the
    /// part each belongs to: an attached template, a linked picture or object, a frame, a
    /// subdocument. Hyperlinks, which are only followed when clicked, are not among them.
    /// </summary>
    public IReadOnlyList<OpcExternalRelationship> ExternalReferences
    {
        get
        {
            List<OpcExternalRelationship> found = [];

            foreach (string source in RelationshipSources())
            {
                foreach (OpcRelationship relationship in relationships[source])
                {
                    if (ActiveContentRules.OfRelationship(relationship) == ActiveContent.ExternalReferences)
                    {
                        found.Add(new OpcExternalRelationship(source, relationship));
                    }
                }
            }

            return found;
        }
    }

    /// <summary>
    /// Takes out everything <see cref="ActiveContent"/> reports, so that the next save writes a
    /// file with none of it: the VBA project (the file becomes a .docx or .dotx), ActiveX
    /// controls, embedded OLE objects and packages, imported chunks, external references other
    /// than hyperlinks, with the elements that referred to them, and fields that start programs
    /// or fetch things, which keep the text Word last showed for them. Pictures that showed an
    /// object or a control stay as pictures. For a file from someone you do not trust.
    /// </summary>
    public void RemoveActiveContent()
    {
        HashSet<string> drop = new(StringComparer.OrdinalIgnoreCase);

        foreach ((string name, PartData part) in parts)
        {
            if (ActiveContentRules.OfPart(name, part.ContentType) != ActiveContent.None)
            {
                drop.Add(name);
            }
        }

        foreach (string source in RelationshipSources())
        {
            foreach (OpcRelationship relationship in relationships[source])
            {
                if (relationship.IsExternal == false && ActiveContentRules.OfRelationship(relationship) != ActiveContent.None)
                {
                    drop.Add(OpcPath.ResolveTarget(source, relationship.Target));
                }
            }
        }

        RemoveMacros();
        Dictionary<string, HashSet<string>> removed = DropParts(drop, removeExternal: true);

        foreach ((string source, HashSet<string> ids) in removed)
        {
            if (ids.Count > 0 && parts.TryGetValue(source, out PartData? part) == true && IsXmlContentType(part.ContentType) == true)
            {
                RemoveReferences(Xml(source).Root, ids);
            }
        }

        foreach ((string name, PartData part) in parts.ToList())
        {
            if (ActiveContentRules.HoldsFields(part.ContentType) == false)
            {
                continue;
            }

            XElement? root = PeekXml(name).Root;

            if (root is not null && FieldScanner.HasActiveField(WordNames.For(root), root) == true)
            {
                XElement? editable = Xml(name).Root;

                if (editable is not null)
                {
                    FieldScanner.Neutralize(WordNames.For(editable), editable);
                }
            }
        }

        contexts.Clear();
    }

    /// <summary>
    /// Removes parts with their own relationships and every relationship that points to them, and,
    /// with <paramref name="removeExternal"/>, every external relationship but hyperlinks. Gives
    /// back the ids removed, per part that held them.
    /// </summary>
    private Dictionary<string, HashSet<string>> DropParts(HashSet<string> drop, bool removeExternal)
    {
        foreach (string name in drop)
        {
            parts.Remove(name);
            relationships.Remove(name);
        }

        Dictionary<string, HashSet<string>> removed = new(StringComparer.OrdinalIgnoreCase);

        foreach ((string source, List<OpcRelationship> list) in relationships)
        {
            HashSet<string> ids = new(StringComparer.Ordinal);

            foreach (OpcRelationship relationship in list)
            {
                bool gone = relationship.IsExternal == false && drop.Contains(OpcPath.ResolveTarget(source, relationship.Target)) == true;
                bool external = removeExternal == true && ActiveContentRules.OfRelationship(relationship) == ActiveContent.ExternalReferences;

                if (gone == true || external == true)
                {
                    ids.Add(relationship.Id);
                }
            }

            list.RemoveAll(relationship => ids.Contains(relationship.Id) == true);
            removed[source] = ids;
        }

        return removed;
    }

    /// <summary>Takes out what referred to removed relationships: whole elements that are only a reference, otherwise the attribute.</summary>
    private static void RemoveReferences(XElement? root, HashSet<string> ids)
    {
        if (root is null)
        {
            return;
        }

        foreach (XElement element in root.DescendantsAndSelf().ToList())
        {
            List<XAttribute> references = [.. element.Attributes().Where(attribute => RelationshipNamespaces.Contains(attribute.Name.NamespaceName) == true && ids.Contains(attribute.Value) == true)];

            if (references.Count == 0)
            {
                continue;
            }

            if (ReferenceOnlyElements.Contains(element.Name.LocalName) == true && element.Parent is not null)
            {
                element.Remove();

                continue;
            }

            foreach (XAttribute reference in references)
            {
                reference.Remove();
            }
        }
    }

    private bool HasActiveFields()
    {
        foreach ((string name, PartData part) in parts)
        {
            if (ActiveContentRules.HoldsFields(part.ContentType) == false)
            {
                continue;
            }

            XElement? root = PeekXml(name).Root;

            if (root is not null && FieldScanner.HasActiveField(WordNames.For(root), root) == true)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>The XML of a part for reading only: the part is not marked as changed, so it is still saved byte for byte.</summary>
    private XDocument PeekXml(string partName)
    {
        PartData part = parts[partName];

        if (part.Xml is not null)
        {
            return part.Xml;
        }

        return WordXml.Load(part.Data, Options.Limits.MaxXmlDepth);
    }

    /// <summary>The parts whose relationships count: the package itself and the parts still in it.</summary>
    private List<string> RelationshipSources()
    {
        return relationships.Keys.Where(source => source.Length == 0 || parts.ContainsKey(source) == true).ToList();
    }

    private static bool IsXmlContentType(string contentType)
    {
        return contentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase) == true || contentType.EndsWith("/xml", StringComparison.OrdinalIgnoreCase) == true;
    }
}
