using System.Globalization;
using System.Security.Cryptography;
using System.Xml.Linq;
using Avig.OpenXml;
using Avig.Ord.Internal;

namespace Avig.Ord;

public sealed partial class OrdDocument
{
    private readonly HashSet<string> usedStyles = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> mediaByHash = new(StringComparer.Ordinal);
    private string? bulletList;
    private string? numberedList;
    private int drawingId = -1;

    /// <summary>A new, empty document with one section on the page given (A4 when null).</summary>
    /// <example>
    /// <code>
    /// OrdDocument document = OrdDocument.Create();
    /// document.AddHeading("Minutes", 1);
    /// document.AddParagraph("Present: …");
    /// await document.SaveAsync("minutes.docx");
    /// </code>
    /// </example>
    public static OrdDocument Create(PageSetup? page = null)
    {
        page ??= PageSetup.A4;
        WordNames names = WordNames.Transitional;
        Dictionary<string, PartData> parts = new(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, List<OpcRelationship>> relationships = new(StringComparer.OrdinalIgnoreCase);
        List<string> order = [];
        OrdDocument document = new(ReadOptions.Default, "word/document.xml", parts, order, relationships);

        XElement root = new(
            names.W + "document",
            new XAttribute(XNamespace.Xmlns + "w", names.W.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "r", names.R.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "wp", WordNames.Drawing),
            new XAttribute(XNamespace.Xmlns + "a", WordNames.DrawingMain),
            new XAttribute(XNamespace.Xmlns + "pic", WordNames.Picture),
            new XElement(names.Body, page.ToSectionProperties(names, [], false)));

        document.AddPart("word/document.xml", WordTypes.DocumentContentType, new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root));
        relationships[string.Empty] = [new OpcRelationship("rId1", OpcNamespaces.OfficeDocumentRelationship, "word/document.xml", false)];
        document.AddPart("word/styles.xml", WordTypes.StylesContentType, BuiltInStyles.NewStylesPart(names));
        document.AddRelationship("word/document.xml", WordTypes.StylesRelationship, "word/styles.xml");
        document.AddPart("word/settings.xml", WordTypes.SettingsContentType, BuiltInStyles.NewSettingsPart(names));
        document.AddRelationship("word/document.xml", WordTypes.SettingsRelationship, "word/settings.xml");

        return document;
    }

    /// <summary>Adds a paragraph of plain text at the end of the body.</summary>
    /// <returns>The paragraph.</returns>
    public Paragraph AddParagraph(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        return AddParagraph(paragraph => paragraph.Text(text));
    }

    /// <summary>Adds a paragraph built with <see cref="ParagraphBuilder"/> at the end of the body.</summary>
    /// <returns>The paragraph.</returns>
    public Paragraph AddParagraph(Action<ParagraphBuilder> build)
    {
        ArgumentNullException.ThrowIfNull(build);
        ParagraphBuilder builder = new();
        build(builder);
        DocumentContext context = MainContext;
        XElement element = builder.Build(SinkFor(context));
        AppendToBody(element);

        return new Paragraph(context, element);
    }

    /// <summary>Adds a heading, level 1 to 9, at the end of the body.</summary>
    /// <returns>The heading's paragraph.</returns>
    public Paragraph AddHeading(string text, int level = 1)
    {
        ArgumentNullException.ThrowIfNull(text);

        return AddParagraph(paragraph => paragraph.Heading(level).Text(text));
    }

    /// <summary>Adds a table built with <see cref="TableBuilder"/> at the end of the body.</summary>
    /// <returns>The table.</returns>
    public Table AddTable(Action<TableBuilder> build)
    {
        ArgumentNullException.ThrowIfNull(build);
        TableBuilder builder = new();
        build(builder);
        DocumentContext context = MainContext;
        XElement element = builder.Build(SinkFor(context), TextWidth());
        AppendToBody(element);

        // Word wants a paragraph after a table; one is added so the next table does not join it.
        AppendToBody(new XElement(context.Names.Paragraph));

        return new Table(context, element);
    }

    /// <summary>Accepts every tracked change in the body, headers, footers and notes: insertions become ordinary text, deletions go.</summary>
    public void AcceptAllRevisions()
    {
        foreach (XElement root in ContentRoots())
        {
            Revisions.Accept(WordNames.For(root), root);
        }

        ChangedEverywhere();
    }

    /// <summary>Rejects every tracked change: insertions go, deletions come back as ordinary text.</summary>
    public void RejectAllRevisions()
    {
        foreach (XElement root in ContentRoots())
        {
            Revisions.Reject(WordNames.For(root), root);
        }

        ChangedEverywhere();
    }

    /// <summary>The builders' view of the document, for content in the part of <paramref name="context"/>.</summary>
    internal IContentSink SinkFor(DocumentContext context)
    {
        return new Sink(this, context);
    }

    /// <summary>The roots of every part with text: the body, the headers and footers, the notes and the comments.</summary>
    internal IEnumerable<XElement> ContentRoots()
    {
        yield return Body;

        foreach (HeaderFooter header in Headers.Concat(Footers))
        {
            yield return header.Root;
        }

        foreach (string kind in (string[])["footnotes", "endnotes", "comments"])
        {
            string? part = RelatedPart(MainPartName, kind);

            if (part is not null && Xml(part).Root is XElement root)
            {
                yield return root;
            }
        }
    }

    internal void ChangedEverywhere()
    {
        foreach (DocumentContext context in contexts.Values)
        {
            context.Changed();
        }
    }

    private void AppendToBody(XElement element)
    {
        XElement body = Body;
        XElement? last = body.Elements().LastOrDefault();

        if (last is not null && last.Name == WordNames.For(body.Document?.Root).SectionProperties)
        {
            last.AddBeforeSelf(element);
        }
        else
        {
            body.Add(element);
        }

        tracking?.InsertBlock(MainContext.Names, element);
        MainContext.Changed();
    }

    /// <summary>The width between the margins of the last section.</summary>
    private Length TextWidth()
    {
        IReadOnlyList<Section> sections = Sections;
        Section? section = null;

        if (sections.Count > 0)
        {
            section = sections[^1];
        }

        Length width = PageSetup.A4.TextWidth;

        if (section?.PageWidth is Length page)
        {
            width = page - (section.MarginLeft ?? Length.Zero) - (section.MarginRight ?? Length.Zero);
        }

        // Margins wider than the page (the schema allows them) gave a table negative column widths.
        if (width <= Length.Zero)
        {
            width = PageSetup.A4.TextWidth;
        }

        return width;
    }

    /// <summary>
    /// The id after the largest one in use, at least <paramref name="minimum"/>; or, when the
    /// largest is int.MaxValue, the smallest one from <paramref name="minimum"/> not in use. A
    /// file may use any id, and one past int.MaxValue became a negative id the schema forbids.
    /// </summary>
    internal static int NextFreeId(IEnumerable<int> used, int minimum)
    {
        HashSet<int> taken = [.. used];
        int largest = taken.DefaultIfEmpty(minimum - 1).Max();

        if (largest < int.MaxValue)
        {
            return Math.Max(largest + 1, minimum);
        }

        int candidate = minimum;

        while (taken.Contains(candidate) == true)
        {
            candidate++;
        }

        return candidate;
    }

    /// <summary>Adds the styles the builders used to the styles part, made when the document has none, and the styles they are based on.</summary>
    partial void BeforeSave()
    {
        if (usedStyles.Count == 0)
        {
            return;
        }

        string stylesPart = RelatedPart(MainPartName, "styles") ?? "word/styles.xml";

        if (HasPart(stylesPart) == false)
        {
            AddPart(stylesPart, WordTypes.StylesContentType, BuiltInStyles.NewStylesPart(MainContext.Names));
            AddRelationship(MainPartName, WordTypes.StylesRelationship, stylesPart);
        }

        XElement root = Xml(stylesPart).Root!;
        WordNames names = WordNames.For(root);
        HashSet<string> present = new(root.Elements(names.Style).Select(style => style.Attribute(names.StyleId)?.Value ?? string.Empty), StringComparer.Ordinal);
        Queue<string> wanted = new(usedStyles);

        while (wanted.Count > 0)
        {
            string id = wanted.Dequeue();

            if (present.Contains(id) == true || BuiltInStyles.TryCreate(names, id, out XElement style) == false)
            {
                continue;
            }

            root.Add(style);
            present.Add(id);
            string? basedOn = names.ValueOf(style, "basedOn");

            if (basedOn is not null)
            {
                wanted.Enqueue(basedOn);
            }
        }

        styles = null;
        ChangedEverywhere();
    }

    /// <summary>The numbering part, made when the document has none.</summary>
    private XElement NumberingRoot()
    {
        string? part = RelatedPart(MainPartName, "numbering");

        if (part is null)
        {
            part = "word/numbering.xml";
            AddPart(part, WordTypes.NumberingContentType, BuiltInStyles.NewNumberingPart(MainContext.Names));
            AddRelationship(MainPartName, WordTypes.NumberingRelationship, part);
        }

        return Xml(part).Root!;
    }

    /// <summary>
    /// A w:num for a new bulleted or numbered list. The list definitions the builders use are
    /// added to the document's numbering part under ids it does not use yet.
    /// </summary>
    private string NewList(bool bullet)
    {
        XElement root = NumberingRoot();
        WordNames names = WordNames.For(root);
        string abstractId = EnsureAbstract(root, names, bullet);
        int next = NextFreeId(root.Elements(names.W + "num").Select(element => FormatReader.Integer(element.Attribute(names.W + "numId")?.Value) ?? 0), 1);
        string id = next.ToString(CultureInfo.InvariantCulture);
        XElement instance = BuiltInStyles.NumberingInstance(names, id, bullet);
        instance.Element(names.W + "abstractNumId")!.SetAttributeValue(names.Val, abstractId);
        root.Add(instance);
        numbering = null;
        ChangedEverywhere();

        return id;
    }

    /// <summary>The id of the builders' bulleted or numbered list definition in the document, added under a new id when it is not there.</summary>
    private static string EnsureAbstract(XElement root, WordNames names, bool bullet)
    {
        string marker = "AvigOrdNumbered";

        if (bullet == true)
        {
            marker = "AvigOrdBullet";
        }

        foreach (XElement existing in root.Elements(names.W + "abstractNum"))
        {
            if (names.ValueOf(existing, "name") == marker)
            {
                return existing.Attribute(names.W + "abstractNumId")?.Value ?? "0";
            }
        }

        int next = NextFreeId(root.Elements(names.W + "abstractNum").Select(element => FormatReader.Integer(element.Attribute(names.W + "abstractNumId")?.Value) ?? 0), 0);
        string template = "1";

        if (bullet == true)
        {
            template = "0";
        }

        XElement definition = new(BuiltInStyles.NewNumberingPart(names).Root!.Elements(names.W + "abstractNum").First(element => element.Attribute(names.W + "abstractNumId")?.Value == template));
        string id = next.ToString(CultureInfo.InvariantCulture);
        definition.SetAttributeValue(names.W + "abstractNumId", id);
        definition.Element(names.W + "multiLevelType")!.AddAfterSelf(RunStyle.Value(names, "name", marker));

        // Abstract definitions come before the w:num elements.
        XElement? firstInstance = root.Elements(names.W + "num").FirstOrDefault();

        if (firstInstance is null)
        {
            root.Add(definition);
        }
        else
        {
            firstInstance.AddBeforeSelf(definition);
        }

        return id;
    }

    /// <summary>The namespace declarations pictures need, on the root of the part, so that the file does not get made-up prefixes.</summary>
    private static void DeclareDrawingNamespaces(XElement root)
    {
        (string Prefix, string Namespace)[] needed = [("wp", WordNames.Drawing), ("a", WordNames.DrawingMain), ("pic", WordNames.Picture)];

        foreach ((string prefix, string ns) in needed)
        {
            if (root.GetPrefixOfNamespace(ns) is null && root.Attribute(XNamespace.Xmlns + prefix) is null)
            {
                root.Add(new XAttribute(XNamespace.Xmlns + prefix, ns));
            }
        }
    }

    /// <summary>The builders' view of the document for one part: relationships, pictures and notes go straight into the document.</summary>
    private sealed class Sink : IContentSink
    {
        private readonly OrdDocument document;
        private readonly DocumentContext context;

        public Sink(OrdDocument document, DocumentContext context)
        {
            context.ThrowIfReadOnly();
            this.document = document;
            this.context = context;
        }

        public WordNames Names => context.Names;

        public Length TextWidth => document.TextWidth();

        public string AddHyperlink(string url)
        {
            return document.AddExternalRelationship(context.PartName, OpcNamespaces.HyperlinkRelationship, url);
        }

        public string AddPicture(byte[] data, string contentType, string extension)
        {
            string hash = Convert.ToHexString(SHA256.HashData(data));

            if (document.mediaByHash.TryGetValue(hash, out string? media) == false)
            {
                media = document.UniquePartName("word/media", "image", extension);
                document.AddPart(media, contentType, data);
                document.mediaByHash[hash] = media;
            }

            if (document.Xml(context.PartName).Root is XElement root)
            {
                DeclareDrawingNamespaces(root);
            }

            string relative = OpcPath.RelativeTarget(context.PartName, media);
            OpcRelationship? existing = document.RelationshipsOf(context.PartName).FirstOrDefault(relationship => relationship.IsExternal == false && relationship.Target == relative);

            if (existing is not null)
            {
                return existing.Id;
            }

            return document.AddRelationship(context.PartName, OpcNamespaces.ImageRelationship, media);
        }

        public int NextDrawingId()
        {
            if (document.drawingId < 0)
            {
                // One before the next free id, so that an id of int.MaxValue in the file does not wrap round.
                document.drawingId = NextFreeId(
                    document.ContentRoots()
                        .SelectMany(root => root.Descendants().Where(element => element.Name.LocalName == "docPr"))
                        .Select(element => FormatReader.Integer(element.Attribute("id")?.Value) ?? 0),
                    1) - 1;
            }

            document.drawingId++;

            return document.drawingId;
        }

        public string AddNote(NoteKind kind, IReadOnlyList<XElement> paragraphs)
        {
            string relationship = "footnotes";
            string part = "word/footnotes.xml";
            string contentType = WordTypes.FootnotesContentType;
            string type = WordTypes.FootnotesRelationship;
            string element = "footnote";

            if (kind == NoteKind.Endnote)
            {
                relationship = "endnotes";
                part = "word/endnotes.xml";
                contentType = WordTypes.EndnotesContentType;
                type = WordTypes.EndnotesRelationship;
                element = "endnote";
            }

            string? existing = document.RelatedPart(document.MainPartName, relationship);

            if (existing is null)
            {
                document.AddPart(part, contentType, BuiltInStyles.NewNotesPart(context.Names, kind));
                document.AddRelationship(document.MainPartName, type, part);
                existing = part;
            }

            XElement root = document.Xml(existing).Root!;
            WordNames names = WordNames.For(root);
            int next = NextFreeId(root.Elements(names.W + element).Select(note => FormatReader.Integer(note.Attribute(names.WId)?.Value) ?? 0), 1);
            string id = next.ToString(CultureInfo.InvariantCulture);
            root.Add(new XElement(names.W + element, new XAttribute(names.WId, id), paragraphs));

            return id;
        }

        public void UseStyle(string styleId)
        {
            document.usedStyles.Add(styleId);
        }

        public string ListNumbering(bool bullet, bool startNew)
        {
            if (bullet == true && document.bulletList is not null && startNew == false)
            {
                return document.bulletList;
            }

            if (bullet == false && document.numberedList is not null && startNew == false)
            {
                return document.numberedList;
            }

            string id = document.NewList(bullet);

            if (bullet == true)
            {
                document.bulletList = id;
            }
            else
            {
                document.numberedList = id;
            }

            return id;
        }
    }
}
