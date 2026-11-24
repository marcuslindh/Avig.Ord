using System.Globalization;
using System.Xml.Linq;
using Avig.OpenXml;
using Avig.Ord.Internal;

namespace Avig.Ord;

/// <summary>Which pages a header or footer is for.</summary>
public enum HeaderFooterKind
{
    /// <summary>Every page, or the odd pages when even pages have their own.</summary>
    Default,

    /// <summary>The first page of the section, when it has its own.</summary>
    First,

    /// <summary>The even pages, when they have their own.</summary>
    Even,
}

/// <summary>A header or a footer.</summary>
public sealed class HeaderFooter
{
    internal HeaderFooter(DocumentContext context, XElement root, HeaderFooterKind kind, bool isHeader)
    {
        Context = context;
        Root = root;
        Kind = kind;
        IsHeader = isHeader;
    }

    /// <summary>Which pages it is for.</summary>
    public HeaderFooterKind Kind { get; }

    /// <summary>True for a header, false for a footer.</summary>
    public bool IsHeader { get; }

    /// <summary>The part it is in, such as word/header1.xml.</summary>
    public string PartName => Context.PartName;

    /// <summary>Its paragraphs and tables.</summary>
    public IReadOnlyList<Block> Blocks => Block.Read(Context, Root);

    /// <summary>Its text.</summary>
    public string Text => TextRenderer.Render(Blocks, TextOptions.Default);

    internal DocumentContext Context { get; }

    internal XElement Root { get; }
}

/// <summary>A footnote or an endnote.</summary>
public sealed class Note
{
    internal Note(DocumentContext context, XElement element, NoteKind kind)
    {
        Context = context;
        Element = element;
        Kind = kind;
    }

    /// <summary>Footnote or endnote.</summary>
    public NoteKind Kind { get; }

    /// <summary>The id the reference in the text has (<see cref="NoteReference.Id"/>).</summary>
    public string Id => Element.Attribute(Context.Names.WId)?.Value ?? string.Empty;

    /// <summary>The note's paragraphs.</summary>
    public IReadOnlyList<Block> Blocks => Block.Read(Context, Element);

    /// <summary>The note's text.</summary>
    public string Text => TextRenderer.Render(Blocks, TextOptions.Default).Trim();

    internal DocumentContext Context { get; }

    internal XElement Element { get; }
}

/// <summary>A comment in the margin.</summary>
public sealed class Comment
{
    internal Comment(DocumentContext context, XElement element)
    {
        Context = context;
        Element = element;
    }

    /// <summary>The comment's id.</summary>
    public string Id => Element.Attribute(Context.Names.WId)?.Value ?? string.Empty;

    /// <summary>Who wrote it.</summary>
    public string Author => Element.Attribute(Context.Names.W + "author")?.Value ?? string.Empty;

    /// <summary>The writer's initials.</summary>
    public string Initials => Element.Attribute(Context.Names.W + "initials")?.Value ?? string.Empty;

    /// <summary>When it was written, if the file says.</summary>
    public DateTimeOffset? Date
    {
        get
        {
            DateTimeOffset? date = null;

            if (DateTimeOffset.TryParse(Element.Attribute(Context.Names.W + "date")?.Value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out DateTimeOffset parsed) == true)
            {
                date = parsed;
            }

            return date;
        }
    }

    /// <summary>The comment's paragraphs.</summary>
    public IReadOnlyList<Block> Blocks => Block.Read(Context, Element);

    /// <summary>The comment's text.</summary>
    public string Text => TextRenderer.Render(Blocks, TextOptions.Default).Trim();

    internal DocumentContext Context { get; }

    internal XElement Element { get; }
}

/// <summary>Which way a page lies.</summary>
public enum Orientation
{
    /// <summary>Taller than wide.</summary>
    Portrait,

    /// <summary>Wider than tall.</summary>
    Landscape,
}

/// <summary>A section's page settings: size, orientation and margins.</summary>
public sealed class Section
{
    internal Section(DocumentContext context, XElement properties)
    {
        Context = context;
        Properties = properties;
    }

    /// <summary>The page's width.</summary>
    public Length? PageWidth => FormatReader.Twips(PageSize?.Attribute(Context.Names.W + "w")?.Value);

    /// <summary>The page's height.</summary>
    public Length? PageHeight => FormatReader.Twips(PageSize?.Attribute(Context.Names.W + "h")?.Value);

    /// <summary>Portrait or landscape.</summary>
    public Orientation Orientation
    {
        get
        {
            if (PageSize?.Attribute(Context.Names.W + "orient")?.Value == "landscape")
            {
                return Orientation.Landscape;
            }

            return Orientation.Portrait;
        }
    }

    /// <summary>The top margin.</summary>
    public Length? MarginTop => Margin("top");

    /// <summary>The bottom margin.</summary>
    public Length? MarginBottom => Margin("bottom");

    /// <summary>The left margin.</summary>
    public Length? MarginLeft => Margin("left");

    /// <summary>The right margin.</summary>
    public Length? MarginRight => Margin("right");

    internal DocumentContext Context { get; }

    internal XElement Properties { get; }

    private XElement? PageSize => Properties.Element(Context.Names.W + "pgSz");

    private Length? Margin(string side)
    {
        return FormatReader.Twips(Properties.Element(Context.Names.W + "pgMar")?.Attribute(Context.Names.W + side)?.Value);
    }
}

/// <summary>The document's properties, as File &gt; Info in Word shows them (docProps/core.xml).</summary>
public sealed class DocumentProperties
{
    private const string CoreNamespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private const string ElementsNamespace = "http://purl.org/dc/elements/1.1/";
    private const string TermsNamespace = "http://purl.org/dc/terms/";

    private readonly Func<XDocument> core;

    internal DocumentProperties(Func<XDocument> core)
    {
        this.core = core;
    }

    /// <summary>The title.</summary>
    public string? Title
    {
        get => Get(ElementsNamespace, "title");
        set => Set(ElementsNamespace, "title", value);
    }

    /// <summary>The subject.</summary>
    public string? Subject
    {
        get => Get(ElementsNamespace, "subject");
        set => Set(ElementsNamespace, "subject", value);
    }

    /// <summary>Who made the document.</summary>
    public string? Author
    {
        get => Get(ElementsNamespace, "creator");
        set => Set(ElementsNamespace, "creator", value);
    }

    /// <summary>The keywords (tags).</summary>
    public string? Keywords
    {
        get => Get(CoreNamespace, "keywords");
        set => Set(CoreNamespace, "keywords", value);
    }

    /// <summary>The description (comments).</summary>
    public string? Description
    {
        get => Get(ElementsNamespace, "description");
        set => Set(ElementsNamespace, "description", value);
    }

    /// <summary>Who saved it last.</summary>
    public string? LastModifiedBy
    {
        get => Get(CoreNamespace, "lastModifiedBy");
        set => Set(CoreNamespace, "lastModifiedBy", value);
    }

    /// <summary>When it was made.</summary>
    public DateTimeOffset? Created
    {
        get => Date("created");
        set => SetDate("created", value);
    }

    /// <summary>When it was last saved.</summary>
    public DateTimeOffset? Modified
    {
        get => Date("modified");
        set => SetDate("modified", value);
    }

    private string? Get(string ns, string name)
    {
        return core().Root?.Element(XName.Get(name, ns))?.Value;
    }

    private void Set(string ns, string name, string? value)
    {
        XElement root = core().Root!;
        XElement? element = root.Element(XName.Get(name, ns));

        if (value is null)
        {
            element?.Remove();

            return;
        }

        if (element is null)
        {
            element = new XElement(XName.Get(name, ns));
            root.Add(element);
        }

        element.Value = value;
    }

    private DateTimeOffset? Date(string name)
    {
        DateTimeOffset? date = null;

        if (DateTimeOffset.TryParse(Get(TermsNamespace, name), CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out DateTimeOffset parsed) == true)
        {
            date = parsed;
        }

        return date;
    }

    private void SetDate(string name, DateTimeOffset? value)
    {
        XElement root = core().Root!;
        XElement? element = root.Element(XName.Get(name, TermsNamespace));

        if (value is null)
        {
            element?.Remove();

            return;
        }

        if (element is null)
        {
            element = new XElement(
                XName.Get(name, TermsNamespace),
                new XAttribute(XName.Get("type", "http://www.w3.org/2001/XMLSchema-instance"), "dcterms:W3CDTF"));
            root.Add(element);
        }

        element.Value = value.Value.UtcDateTime.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
    }
}

public sealed partial class OrdDocument
{
    private DocumentProperties? properties;

    /// <summary>The sections, in order: their page settings. The last is the body's own.</summary>
    public IReadOnlyList<Section> Sections
    {
        get
        {
            DocumentContext context = MainContext;
            WordNames names = context.Names;
            List<Section> sections = [];

            foreach (XElement properties in Body.Descendants(names.SectionProperties))
            {
                sections.Add(new Section(context, properties));
            }

            return sections;
        }
    }

    /// <summary>The headers, in the order the sections refer to them.</summary>
    public IReadOnlyList<HeaderFooter> Headers => HeadersAndFooters("headerReference", true);

    /// <summary>The footers, in the order the sections refer to them.</summary>
    public IReadOnlyList<HeaderFooter> Footers => HeadersAndFooters("footerReference", false);

    /// <summary>The footnotes, without the separators Word keeps among them.</summary>
    public IReadOnlyList<Note> Footnotes => Notes("footnotes", "footnote", NoteKind.Footnote);

    /// <summary>The endnotes, without the separators Word keeps among them.</summary>
    public IReadOnlyList<Note> Endnotes => Notes("endnotes", "endnote", NoteKind.Endnote);

    /// <summary>The comments.</summary>
    public IReadOnlyList<Comment> Comments
    {
        get
        {
            string? part = RelatedPart(MainPartName, "comments");

            if (part is null)
            {
                return [];
            }

            DocumentContext context = ContextFor(part);

            return [.. Xml(part).Root?.Elements(context.Names.W + "comment").Select(element => new Comment(context, element)) ?? []];
        }
    }

    /// <summary>The title, author and the other properties Word shows under File &gt; Info.</summary>
    public DocumentProperties Properties => properties ??= new DocumentProperties(CoreProperties);

    /// <summary>
    /// The document as plain text: each paragraph on a line of its own, list items with their
    /// labels, table rows on lines of their own with the cells separated by tabs.
    /// </summary>
    public string ToPlainText()
    {
        return Text;
    }

    /// <summary>The document as GitHub Flavored Markdown (see <see cref="MarkdownOptions"/>).</summary>
    public string ToMarkdown(MarkdownOptions? options = null)
    {
        return MarkdownRenderer.Render(Blocks, options ?? MarkdownOptions.Default);
    }

    /// <summary>
    /// Writes the document as Markdown to <paramref name="markdownPath"/>, and its pictures to
    /// <see cref="MarkdownOptions.PictureFolder"/> beside it, so that the links work.
    /// </summary>
    /// <returns>Success, or the failure to write a file.</returns>
    public async Task<Result> SaveMarkdownAsync(string markdownPath, MarkdownOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(markdownPath);
        options ??= MarkdownOptions.Default;
        string folder = Path.GetDirectoryName(Path.GetFullPath(markdownPath)) ?? ".";

        try
        {
            await File.WriteAllTextAsync(markdownPath, ToMarkdown(options), cancellationToken);

            if (options.PictureLink is not null)
            {
                return Result.Success();
            }

            string pictures = Path.Combine(folder, options.PictureFolder);

            foreach (Picture picture in AllPictures())
            {
                Result<byte[]> data = await picture.ReadAsync(cancellationToken);

                if (data.IsSuccess(out byte[]? bytes) == true)
                {
                    Directory.CreateDirectory(pictures);
                    await File.WriteAllBytesAsync(Path.Combine(pictures, MarkdownRenderer.PictureFileName(picture)), bytes, cancellationToken);
                }
            }
        }
        catch (UnauthorizedAccessException)
        {
            return Result.Forbidden($"'{markdownPath}' or its picture folder may not be written.");
        }
        catch (IOException exception)
        {
            return Result.Conflict($"'{markdownPath}' could not be written: {exception.Message}");
        }

        return Result.Success();
    }

    /// <summary>Every picture in the body, also in tables and text boxes.</summary>
    internal IEnumerable<Picture> AllPictures()
    {
        foreach (Paragraph paragraph in Paragraphs)
        {
            foreach (Inline inline in InlineReader.Flatten(paragraph.Inlines))
            {
                if (inline is Picture picture)
                {
                    yield return picture;
                }
                else if (inline is TextBox box)
                {
                    foreach (Paragraph inner in AllParagraphs(box.Blocks))
                    {
                        foreach (Picture boxed in InlineReader.Flatten(inner.Inlines).OfType<Picture>())
                        {
                            yield return boxed;
                        }
                    }
                }
            }
        }
    }

    /// <summary>The part the first relationship of a kind points to from <paramref name="source"/>, if it is in the package.</summary>
    internal string? RelatedPart(string source, string kind)
    {
        string? found = null;

        foreach (OpcRelationship relationship in RelationshipsOf(source))
        {
            if (relationship.IsExternal == false && OpcPath.IsRelationship(relationship.Type, kind) == true)
            {
                string target = OpcPath.ResolveTarget(source, relationship.Target);

                if (HasPart(target) == true)
                {
                    found = target;

                    break;
                }
            }
        }

        return found;
    }

    private List<HeaderFooter> HeadersAndFooters(string reference, bool header)
    {
        DocumentContext main = MainContext;
        WordNames names = main.Names;
        List<HeaderFooter> list = [];
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);

        foreach (XElement element in Body.Descendants(names.W + reference))
        {
            string? id = element.Attribute(names.Id)?.Value;

            if (id is null || main.TryResolve(id, out string part, out bool external) == false || external == true || HasPart(part) == false || seen.Add(part) == false)
            {
                continue;
            }

            HeaderFooterKind kind = element.Attribute(names.Type)?.Value switch
            {
                "first" => HeaderFooterKind.First,
                "even" => HeaderFooterKind.Even,
                _ => HeaderFooterKind.Default,
            };

            DocumentContext context = ContextFor(part);
            XElement? root = Xml(part).Root;

            if (root is not null)
            {
                list.Add(new HeaderFooter(context, root, kind, header));
            }
        }

        return list;
    }

    private IReadOnlyList<Note> Notes(string relationship, string element, NoteKind kind)
    {
        string? part = RelatedPart(MainPartName, relationship);

        if (part is null)
        {
            return [];
        }

        DocumentContext context = ContextFor(part);
        WordNames names = context.Names;

        return [.. Xml(part).Root?.Elements(names.W + element)
            .Where(note => note.Attribute(names.Type)?.Value is null or "normal")
            .Select(note => new Note(context, note, kind)) ?? []];
    }

    /// <summary>docProps/core.xml, made when the document has none.</summary>
    private XDocument CoreProperties()
    {
        foreach (OpcRelationship relationship in RelationshipsOf(string.Empty))
        {
            if (OpcPath.IsRelationship(relationship.Type, "core-properties") == true)
            {
                string part = OpcPath.ResolveTarget(string.Empty, relationship.Target);

                if (HasPart(part) == true)
                {
                    return Xml(part);
                }
            }
        }

        const string name = "docProps/core.xml";
        XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        XDocument created = new(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(
                cp + "coreProperties",
                new XAttribute(XNamespace.Xmlns + "cp", cp.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "dc", "http://purl.org/dc/elements/1.1/"),
                new XAttribute(XNamespace.Xmlns + "dcterms", "http://purl.org/dc/terms/"),
                new XAttribute(XNamespace.Xmlns + "xsi", "http://www.w3.org/2001/XMLSchema-instance")));

        AddPart(name, OpcNamespaces.CorePropertiesContentType, created);
        AddRelationship(string.Empty, OpcNamespaces.CorePropertiesRelationship, name);

        return created;
    }
}
