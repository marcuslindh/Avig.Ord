using System.Text;
using System.Xml;
using System.Xml.Linq;
using Avig.OpenXml;

namespace Avig.Ord.Internal;

/// <summary>The content types and relationship kinds of WordprocessingML.</summary>
internal static class WordTypes
{
    public const string DocumentContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    public const string TemplateContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";
    public const string MacroDocumentContentType = "application/vnd.ms-word.document.macroEnabled.main+xml";
    public const string MacroTemplateContentType = "application/vnd.ms-word.template.macroEnabledTemplate.main+xml";
    public const string StylesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    public const string NumberingContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
    public const string SettingsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    public const string HeaderContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
    public const string FooterContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
    public const string FootnotesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
    public const string EndnotesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
    public const string CommentsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

    public const string StylesRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    public const string NumberingRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    public const string SettingsRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    public const string HeaderRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    public const string FooterRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    public const string FootnotesRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    public const string EndnotesRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    public const string CommentsRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

    public static bool IsMainContentType(string? contentType)
    {
        return contentType is DocumentContentType or TemplateContentType or MacroDocumentContentType or MacroTemplateContentType;
    }

    /// <summary>The content type a picture's bytes have, from their first bytes; application/octet-stream when they are none of the formats Word shows.</summary>
    public static string ImageContentType(ReadOnlySpan<byte> data, out string extension)
    {
        extension = "bin";

        if (data.StartsWith((ReadOnlySpan<byte>)[0x89, 0x50, 0x4E, 0x47]) == true)
        {
            extension = "png";

            return "image/png";
        }

        if (data.StartsWith((ReadOnlySpan<byte>)[0xFF, 0xD8, 0xFF]) == true)
        {
            extension = "jpeg";

            return "image/jpeg";
        }

        if (data.StartsWith("GIF8"u8) == true)
        {
            extension = "gif";

            return "image/gif";
        }

        if (data.StartsWith("BM"u8) == true)
        {
            extension = "bmp";

            return "image/bmp";
        }

        ReadOnlySpan<byte> head = data[..Math.Min(data.Length, 512)];

        if (head.IndexOf("<svg"u8) >= 0)
        {
            extension = "svg";

            return "image/svg+xml";
        }

        return "application/octet-stream";
    }
}

/// <summary>Loads and saves the XML parts of a document as <see cref="XDocument"/>s.</summary>
internal static class WordXml
{
    /// <summary>Loads a part, with the limits of the file (no DTD, a depth limit).</summary>
    public static XDocument Load(byte[] data, int maxDepth)
    {
        using MemoryStream stream = new(data, false);
        using XmlReader reader = OpcXml.CreateReader(stream, maxDepth);

        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    public static byte[] Save(XDocument document)
    {
        using MemoryStream buffer = new();
        XmlWriterSettings settings = new() { Encoding = new UTF8Encoding(false), Indent = false };

        using (XmlWriter writer = XmlWriter.Create(buffer, settings))
        {
            document.Declaration ??= new XDeclaration("1.0", "UTF-8", "yes");
            document.Save(writer);
        }

        return buffer.ToArray();
    }
}
