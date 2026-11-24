using System.Globalization;
using System.Xml.Linq;
using Avig.OpenXml;
using Avig.Ord.Internal;

namespace Avig.Ord;

public sealed partial class OrdDocument
{
    /// <summary>Adds a new XML part. It is saved from its XML.</summary>
    internal void AddPart(string name, string contentType, XDocument xml)
    {
        parts[name] = new PartData(contentType, []) { Xml = xml, Dirty = true };

        if (partOrder.Contains(name, StringComparer.OrdinalIgnoreCase) == false)
        {
            partOrder.Add(name);
        }
    }

    /// <summary>Adds a new part from bytes, such as a picture.</summary>
    internal void AddPart(string name, string contentType, byte[] data)
    {
        parts[name] = new PartData(contentType, data);

        if (partOrder.Contains(name, StringComparer.OrdinalIgnoreCase) == false)
        {
            partOrder.Add(name);
        }
    }

    /// <summary>A part name not yet in the package: word/media/image1.png, image2.png …</summary>
    internal string UniquePartName(string folder, string stem, string extension)
    {
        for (int number = 1; ; number++)
        {
            string name = $"{folder}/{stem}{number.ToString(CultureInfo.InvariantCulture)}.{extension}";

            if (parts.ContainsKey(name) == false)
            {
                return name;
            }
        }
    }

    /// <summary>Adds a relationship from a part to another and gives back its id, one the part does not use yet.</summary>
    internal string AddRelationship(string source, string type, string targetPart)
    {
        List<OpcRelationship> list = (List<OpcRelationship>)RelationshipsOf(source);
        string id = NextRelationshipId(list);
        string target = targetPart;

        if (source.Length > 0)
        {
            target = OpcPath.RelativeTarget(source, targetPart);
        }

        list.Add(new OpcRelationship(id, type, target, false));

        return id;
    }

    /// <summary>Adds a relationship to a URL (a hyperlink) and gives back its id; an existing one to the same URL is used again.</summary>
    internal string AddExternalRelationship(string source, string type, string url)
    {
        List<OpcRelationship> list = (List<OpcRelationship>)RelationshipsOf(source);
        OpcRelationship? existing = list.FirstOrDefault(relationship => relationship.IsExternal == true && relationship.Type == type && relationship.Target == url);

        if (existing is not null)
        {
            return existing.Id;
        }

        string id = NextRelationshipId(list);
        list.Add(new OpcRelationship(id, type, url, true));

        return id;
    }

    private static string NextRelationshipId(List<OpcRelationship> list)
    {
        HashSet<string> used = new(list.Select(relationship => relationship.Id), StringComparer.Ordinal);

        for (int number = list.Count + 1; ; number++)
        {
            string id = "rId" + number.ToString(CultureInfo.InvariantCulture);

            if (used.Contains(id) == false)
            {
                return id;
            }
        }
    }
}
