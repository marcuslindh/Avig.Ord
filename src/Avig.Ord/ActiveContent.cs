using Avig.OpenXml;

namespace Avig.Ord;

/// <summary>
/// What in a Word file can act on its own when Word opens it, or fetch something from
/// elsewhere: the ways malware and phishing get into documents. The package never runs, follows
/// or updates any of it, but it keeps it when a document is saved, so a file passed on is no
/// safer for having been opened here. <see cref="OrdDocument.RemoveActiveContent"/> takes it out.
/// </summary>
[Flags]
public enum ActiveContent
{
    /// <summary>None of it.</summary>
    None = 0,

    /// <summary>A VBA project (a .docm or .dotm).</summary>
    Macros = 1,

    /// <summary>ActiveX controls, which run code when the document is opened if Word lets them.</summary>
    ActiveX = 2,

    /// <summary>
    /// Embedded OLE objects and packages (an Excel sheet, a program, any file shown as an icon),
    /// and imported documents (w:altChunk: HTML, RTF or MHT that Word converts when it opens the
    /// file). OLE is where Equation Editor exploits and CVE-2025-21298 live.
    /// </summary>
    EmbeddedObjects = 4,

    /// <summary>
    /// Fields that start a program or fetch something when they are updated: DDE and DDEAUTO,
    /// INCLUDETEXT, INCLUDEPICTURE, LINK, IMPORT and MACROBUTTON.
    /// </summary>
    ActiveFields = 8,

    /// <summary>
    /// Relationships to something outside the file that Word follows on its own, unlike
    /// hyperlinks: a template attached from a server (remote template injection), a linked
    /// picture or OLE object, a frame, a subdocument (see <see cref="OrdDocument.ExternalReferences"/>).
    /// </summary>
    ExternalReferences = 16,
}

/// <summary>How a part, a relationship or a field instruction counts as active content; one place for the document model and the streaming reader.</summary>
internal static class ActiveContentRules
{
    private static readonly string[] ActiveFieldNames = ["DDE", "DDEAUTO", "INCLUDETEXT", "INCLUDEPICTURE", "LINK", "IMPORT", "MACROBUTTON"];

    /// <summary>What a part is, from its name and content type.</summary>
    public static ActiveContent OfPart(string name, string? contentType)
    {
        string type = contentType ?? string.Empty;

        if (type.StartsWith("application/vnd.ms-office.vba", StringComparison.OrdinalIgnoreCase) == true || name.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase) == true)
        {
            return ActiveContent.Macros;
        }

        if (string.Equals(type, OpcNamespaces.ActiveXContentType, StringComparison.OrdinalIgnoreCase) == true || string.Equals(type, OpcNamespaces.ActiveXBinaryContentType, StringComparison.OrdinalIgnoreCase) == true)
        {
            return ActiveContent.ActiveX;
        }

        if (string.Equals(type, OpcNamespaces.OleObjectContentType, StringComparison.OrdinalIgnoreCase) == true)
        {
            return ActiveContent.EmbeddedObjects;
        }

        return ActiveContent.None;
    }

    /// <summary>What a relationship is: one to a control, an embedded object or package, an imported chunk, or an external target Word follows.</summary>
    public static ActiveContent OfRelationship(OpcRelationship relationship)
    {
        if (relationship.IsExternal == true)
        {
            if (OpcPath.IsRelationship(relationship.Type, "hyperlink") == true)
            {
                return ActiveContent.None;
            }

            return ActiveContent.ExternalReferences;
        }

        if (OpcPath.IsRelationship(relationship.Type, "control") == true || OpcPath.IsRelationship(relationship.Type, "activeXControlBinary") == true)
        {
            return ActiveContent.ActiveX;
        }

        if (OpcPath.IsRelationship(relationship.Type, "oleObject") == true || OpcPath.IsRelationship(relationship.Type, "package") == true || OpcPath.IsRelationship(relationship.Type, "aFChunk") == true)
        {
            return ActiveContent.EmbeddedObjects;
        }

        if (OpcPath.IsRelationship(relationship.Type, "vbaProject") == true)
        {
            return ActiveContent.Macros;
        }

        return ActiveContent.None;
    }

    /// <summary>True for a field instruction whose field starts a program or fetches something: its first word is DDE, INCLUDEPICTURE …</summary>
    public static bool IsActiveField(string instruction)
    {
        string trimmed = instruction.TrimStart();
        int end = 0;

        while (end < trimmed.Length && char.IsWhiteSpace(trimmed[end]) == false && trimmed[end] != '\\' && trimmed[end] != '"')
        {
            end++;
        }

        string name = trimmed[..end];

        return ActiveFieldNames.Contains(name, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>True for a part whose XML can hold fields: the body, headers, footers, notes, comments and the glossary.</summary>
    public static bool HoldsFields(string? contentType)
    {
        if (contentType is null || contentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase) == false)
        {
            return false;
        }

        // The macro-enabled main parts (.docm, .dotm) have types of their own, application/vnd.ms-word.
        return contentType.Contains("wordprocessingml", StringComparison.OrdinalIgnoreCase) == true || contentType.StartsWith("application/vnd.ms-word.", StringComparison.OrdinalIgnoreCase) == true;
    }
}
