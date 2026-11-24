using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Avig.OpenXml;
using Avig.Ord.Internal;

namespace Avig.Ord;

/// <summary>A section's page: its size, orientation and margins.</summary>
public sealed record PageSetup
{
    /// <summary>A4 portrait with 2.5 cm margins, as Swedish Word starts a document.</summary>
    public static PageSetup A4 { get; } = new();

    /// <summary>US Letter portrait with 1 inch margins.</summary>
    public static PageSetup Letter { get; } = new()
    {
        Width = Length.Inch(8.5),
        Height = Length.Inch(11),
        MarginTop = Length.Inch(1),
        MarginBottom = Length.Inch(1),
        MarginLeft = Length.Inch(1),
        MarginRight = Length.Inch(1),
    };

    /// <summary>The page's width.</summary>
    public Length Width { get; init; } = Length.FromEmu(11_906 * Length.EmuPerTwip);

    /// <summary>The page's height.</summary>
    public Length Height { get; init; } = Length.FromEmu(16_838 * Length.EmuPerTwip);

    /// <summary>The top margin.</summary>
    public Length MarginTop { get; init; } = Length.Cm(2.5);

    /// <summary>The bottom margin.</summary>
    public Length MarginBottom { get; init; } = Length.Cm(2.5);

    /// <summary>The left margin.</summary>
    public Length MarginLeft { get; init; } = Length.Cm(2.5);

    /// <summary>The right margin.</summary>
    public Length MarginRight { get; init; } = Length.Cm(2.5);

    /// <summary>How far from the page's top edge the header is.</summary>
    public Length HeaderDistance { get; init; } = Length.Cm(1.25);

    /// <summary>How far from the page's bottom edge the footer is.</summary>
    public Length FooterDistance { get; init; } = Length.Cm(1.25);

    /// <summary>True when the page is wider than tall.</summary>
    public bool IsLandscape => Width > Height;

    /// <summary>The width between the margins.</summary>
    public Length TextWidth => Width - MarginLeft - MarginRight;

    /// <summary>The same page turned: landscape for a portrait one, portrait for a landscape one.</summary>
    public PageSetup Turned()
    {
        return this with { Width = Height, Height = Width };
    }

    internal XElement ToSectionProperties(WordNames names, IEnumerable<XElement> references, bool titlePage)
    {
        XNamespace wml = names.W;
        XElement size = new(wml + "pgSz", new XAttribute(wml + "w", Twips(Width)), new XAttribute(wml + "h", Twips(Height)));

        if (IsLandscape == true)
        {
            size.Add(new XAttribute(wml + "orient", "landscape"));
        }

        XElement properties = new(names.SectionProperties, references);
        properties.Add(
            size,
            new XElement(
                wml + "pgMar",
                new XAttribute(wml + "top", Twips(MarginTop)),
                new XAttribute(wml + "right", Twips(MarginRight)),
                new XAttribute(wml + "bottom", Twips(MarginBottom)),
                new XAttribute(wml + "left", Twips(MarginLeft)),
                new XAttribute(wml + "header", Twips(HeaderDistance)),
                new XAttribute(wml + "footer", Twips(FooterDistance)),
                new XAttribute(wml + "gutter", "0")),
            new XElement(wml + "cols", new XAttribute(wml + "space", "708")));

        if (titlePage == true)
        {
            properties.Add(new XElement(wml + "titlePg"));
        }

        properties.Add(new XElement(wml + "docGrid", new XAttribute(wml + "linePitch", "360")));

        return properties;
    }

    private static string Twips(Length length)
    {
        return length.ToTwips().ToString(CultureInfo.InvariantCulture);
    }
}

/// <summary>How <see cref="DocxWriter"/> writes a document.</summary>
public sealed record DocxWriterOptions
{
    /// <summary>The options used when none are given.</summary>
    public static DocxWriterOptions Default { get; } = new();

    /// <summary>The first section's page. A4 by default.</summary>
    public PageSetup Page { get; init; } = PageSetup.A4;

    /// <summary>A template (.dotx) instead of a document (.docx).</summary>
    public bool Template { get; init; }

    /// <summary>Write Flat OPC (one XML file) instead of a zip. The whole document is then held in memory until it is complete.</summary>
    public bool FlatOpc { get; init; }

    /// <summary>The title, under File &gt; Info.</summary>
    public string? Title { get; init; }

    /// <summary>The author, under File &gt; Info.</summary>
    public string? Author { get; init; }
}

/// <summary>Builds the content of a header or a footer.</summary>
public sealed class HeaderFooterBuilder
{
    private readonly List<Func<IContentSink, Length, XElement>> blocks = [];

    /// <summary>A paragraph of plain text.</summary>
    public HeaderFooterBuilder Paragraph(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        return Paragraph(paragraph => paragraph.Text(text));
    }

    /// <summary>A paragraph built with <see cref="ParagraphBuilder"/>, for example with <see cref="ParagraphBuilder.PageNumber"/>.</summary>
    public HeaderFooterBuilder Paragraph(Action<ParagraphBuilder> build)
    {
        ArgumentNullException.ThrowIfNull(build);
        ParagraphBuilder paragraph = new();
        build(paragraph);
        blocks.Add((sink, _) => paragraph.Build(sink));

        return this;
    }

    /// <summary>A table.</summary>
    public HeaderFooterBuilder Table(Action<TableBuilder> build)
    {
        ArgumentNullException.ThrowIfNull(build);
        TableBuilder table = new();
        build(table);
        blocks.Add((sink, width) => table.Build(sink, width));

        return this;
    }

    internal IEnumerable<XElement> Build(IContentSink sink)
    {
        List<XElement> built = [.. blocks.Select(block => block(sink, sink.TextWidth))];

        if (built.Count == 0 || built[^1].Name != sink.Names.Paragraph)
        {
            built.Add(new XElement(sink.Names.Paragraph));
        }

        return built;
    }
}

/// <summary>
/// Writes a new document block by block, straight to the file: memory stays at one block plus
/// the pictures, which are written when the document is complete. Headings, paragraphs, lists,
/// tables, pictures, links, notes, sections, headers and footers. For a document to build in
/// memory and change before it is saved, use <see cref="OrdDocument.Create(PageSetup?)"/>.
/// </summary>
/// <example>
/// <code>
/// await using DocxWriter writer = await DocxWriter.CreateAsync(stream);
/// await writer.HeadingAsync("Report", 1);
/// await writer.ParagraphAsync(paragraph => paragraph.Text("Total: ").Text("1 200 kr", style => style.Bold()));
/// await writer.TableAsync(t => t.HeaderRow("Item", "Cost").Row("Rent", "8 000"));
/// await writer.CompleteAsync();
/// </code>
/// </example>
public sealed class DocxWriter : IAsyncDisposable
{
    private const string MainPart = "word/document.xml";

    private readonly OpcPackageWriter package;
    private readonly Stream body;
    private readonly XmlWriter xml;
    private readonly DocxWriterOptions options;
    private readonly WordNames names = WordNames.Transitional;
    private readonly HashSet<string> usedStyles = new(StringComparer.Ordinal);
    private readonly List<(string Name, string ContentType, byte[] Data)> media = [];
    private readonly Dictionary<string, string> mediaByHash = new(StringComparer.Ordinal);
    private readonly List<(string Part, string ContentType, XDocument Xml)> extraParts = [];
    private readonly Dictionary<string, List<OpcRelationship>> relationships = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<XElement> numberingInstances = [];
    private readonly List<XElement> footnotes = [];
    private readonly List<XElement> endnotes = [];
    private readonly List<XElement> headerReferences = [];
    private readonly Sink mainSink;
    private PageSetup page;
    private bool titlePage;
    private bool lastWasTable;
    private bool completed;
    private int drawingId;
    private int headerCount;
    private int footerCount;
    private string? bulletList;
    private string? numberedList;

    private DocxWriter(OpcPackageWriter package, Stream body, XmlWriter xml, DocxWriterOptions options)
    {
        this.package = package;
        this.body = body;
        this.xml = xml;
        this.options = options;
        page = options.Page;
        mainSink = new Sink(this, MainPart);
    }

    /// <summary>Starts a document on a stream.</summary>
    /// <param name="output">Where the file goes.</param>
    /// <param name="options">Page, template or document, Flat OPC; <see cref="DocxWriterOptions.Default"/> when null.</param>
    /// <param name="leaveOpen">Leave the stream open when the writer is disposed.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>The writer.</returns>
    public static async Task<DocxWriter> CreateAsync(Stream output, DocxWriterOptions? options = null, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(output);
        options ??= DocxWriterOptions.Default;
        OpcPackageKind kind = OpcPackageKind.Zip;

        if (options.FlatOpc == true)
        {
            kind = OpcPackageKind.FlatOpc;
        }

        OpcPackageWriter package = await OpcPackageWriter.CreateAsync(output, kind, leaveOpen, cancellationToken);
        package.FlatOpcProgId = "Word.Document";
        string contentType = WordTypes.DocumentContentType;

        if (options.Template == true)
        {
            contentType = WordTypes.TemplateContentType;
        }

        Stream body = await package.BeginPartAsync(MainPart, contentType, cancellationToken);
        XmlWriter xml = XmlWriter.Create(body, new XmlWriterSettings { Async = true, Encoding = new UTF8Encoding(false), CloseOutput = false });
        WordNames names = WordNames.Transitional;

        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync("w", "document", names.W.NamespaceName);
        await xml.WriteAttributeStringAsync("xmlns", "w", null, names.W.NamespaceName);
        await xml.WriteAttributeStringAsync("xmlns", "r", null, names.R.NamespaceName);
        await xml.WriteAttributeStringAsync("xmlns", "wp", null, WordNames.Drawing);
        await xml.WriteAttributeStringAsync("xmlns", "a", null, WordNames.DrawingMain);
        await xml.WriteAttributeStringAsync("xmlns", "pic", null, WordNames.Picture);
        await xml.WriteStartElementAsync("w", "body", names.W.NamespaceName);

        return new DocxWriter(package, body, xml, options);
    }

    /// <summary>Starts a document in a file, replacing it if it exists.</summary>
    /// <returns>The writer; NotFound when the folder does not exist, Forbidden or Conflict when the file cannot be written.</returns>
    public static async Task<Result<DocxWriter>> CreateAsync(string path, DocxWriterOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        FileStream file;

        try
        {
            file = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 65_536, FileOptions.Asynchronous);
        }
        catch (DirectoryNotFoundException)
        {
            return Result<DocxWriter>.NotFound($"The folder for '{path}' does not exist.");
        }
        catch (UnauthorizedAccessException)
        {
            return Result<DocxWriter>.Forbidden($"'{path}' may not be written.");
        }
        catch (IOException exception)
        {
            return Result<DocxWriter>.Conflict($"'{path}' could not be written: {exception.Message}");
        }

        return await CreateAsync(file, options, false, cancellationToken);
    }

    /// <summary>A paragraph of plain text.</summary>
    public Task ParagraphAsync(string text, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(text);

        return ParagraphAsync(paragraph => paragraph.Text(text), cancellationToken);
    }

    /// <summary>A paragraph built with <see cref="ParagraphBuilder"/>.</summary>
    public Task ParagraphAsync(Action<ParagraphBuilder> build, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(build);
        ParagraphBuilder paragraph = new();
        build(paragraph);

        return WriteBlockAsync(paragraph.Build(mainSink), false, cancellationToken);
    }

    /// <summary>A heading, level 1 to 9.</summary>
    public Task HeadingAsync(string text, int level = 1, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(text);

        return ParagraphAsync(paragraph => paragraph.Heading(level).Text(text), cancellationToken);
    }

    /// <summary>An item of a bulleted list.</summary>
    public Task BulletAsync(string text, int level = 0, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(text);

        return ParagraphAsync(paragraph => paragraph.Bullet(level).Text(text), cancellationToken);
    }

    /// <summary>An item of a numbered list; <paramref name="startNew"/> starts a new list at 1.</summary>
    public Task NumberedAsync(string text, int level = 0, bool startNew = false, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(text);

        return ParagraphAsync(paragraph => paragraph.Numbered(level, startNew).Text(text), cancellationToken);
    }

    /// <summary>A table built with <see cref="TableBuilder"/>.</summary>
    public Task TableAsync(Action<TableBuilder> build, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(build);
        TableBuilder table = new();
        build(table);

        return WriteBlockAsync(table.Build(mainSink, page.TextWidth), true, cancellationToken);
    }

    /// <summary>A paragraph with a picture at its own size (no wider than the text).</summary>
    public Task PictureAsync(byte[] data, string description = "", CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(data);

        return ParagraphAsync(paragraph => paragraph.Picture(data, description), cancellationToken);
    }

    /// <summary>A page break: what comes next starts on a new page.</summary>
    public Task PageBreakAsync(CancellationToken cancellationToken = default)
    {
        return ParagraphAsync(paragraph => paragraph.PageBreak(), cancellationToken);
    }

    /// <summary>
    /// Ends the section and starts a new one on a new page with another page setup, such as
    /// <c>PageSetup.A4.Turned()</c> for landscape pages. Headers and footers carry on unless
    /// new ones are set.
    /// </summary>
    public async Task SectionBreakAsync(PageSetup next, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(next);
        XElement paragraph = new(names.Paragraph, new XElement(names.ParagraphProperties, page.ToSectionProperties(names, headerReferences, titlePage)));
        await WriteBlockAsync(paragraph, false, cancellationToken);
        headerReferences.Clear();
        titlePage = false;
        page = next;
    }

    /// <summary>The header of the section being written, from here on. <see cref="HeaderFooterKind.First"/> gives the section's first page a header of its own.</summary>
    public void Header(Action<HeaderFooterBuilder> build, HeaderFooterKind kind = HeaderFooterKind.Default)
    {
        HeaderOrFooter(build, kind, true);
    }

    /// <summary>The footer of the section being written. <see cref="HeaderFooterKind.First"/> gives the section's first page a footer of its own.</summary>
    public void Footer(Action<HeaderFooterBuilder> build, HeaderFooterKind kind = HeaderFooterKind.Default)
    {
        HeaderOrFooter(build, kind, false);
    }

    /// <summary>Ends the body and writes the styles, lists, notes, headers, footers and pictures. Nothing can be written afterwards.</summary>
    public async Task CompleteAsync(CancellationToken cancellationToken = default)
    {
        ThrowIfCompleted();
        completed = true;

        if (lastWasTable == true)
        {
            await new XElement(names.Paragraph).WriteToAsync(xml, cancellationToken);
        }

        await page.ToSectionProperties(names, headerReferences, titlePage).WriteToAsync(xml, cancellationToken);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
        await xml.FlushAsync();
        await xml.DisposeAsync();
        await body.DisposeAsync();

        await WritePartsAsync(cancellationToken);
        await package.CompleteAsync(cancellationToken);
    }

    /// <summary>Completes the document if <see cref="CompleteAsync"/> was not called, and closes the file.</summary>
    /// <returns>A task that completes when the file is closed.</returns>
    public async ValueTask DisposeAsync()
    {
        if (completed == false)
        {
            await CompleteAsync();
        }

        await package.DisposeAsync();
    }

    private static XDocument Part(XElement root)
    {
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
    }

    private void HeaderOrFooter(Action<HeaderFooterBuilder> build, HeaderFooterKind kind, bool header)
    {
        ArgumentNullException.ThrowIfNull(build);
        ThrowIfCompleted();
        HeaderFooterBuilder content = new();
        build(content);
        string part;
        string element = "ftr";
        string reference = "footerReference";
        string contentType = WordTypes.FooterContentType;
        string relationship = WordTypes.FooterRelationship;

        if (header == true)
        {
            headerCount++;
            part = "word/header" + headerCount.ToString(CultureInfo.InvariantCulture) + ".xml";
            element = "hdr";
            reference = "headerReference";
            contentType = WordTypes.HeaderContentType;
            relationship = WordTypes.HeaderRelationship;
        }
        else
        {
            footerCount++;
            part = "word/footer" + footerCount.ToString(CultureInfo.InvariantCulture) + ".xml";
        }

        Sink sink = new(this, part);
        XElement root = new(
            names.W + element,
            new XAttribute(XNamespace.Xmlns + "w", names.W.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "r", names.R.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "wp", WordNames.Drawing),
            new XAttribute(XNamespace.Xmlns + "a", WordNames.DrawingMain),
            new XAttribute(XNamespace.Xmlns + "pic", WordNames.Picture),
            content.Build(sink));

        extraParts.Add((part, contentType, Part(root)));
        string id = AddRelationship(MainPart, relationship, part);
        string type = "default";

        if (kind == HeaderFooterKind.First)
        {
            type = "first";
            titlePage = true;
        }
        else if (kind == HeaderFooterKind.Even)
        {
            type = "even";
        }

        headerReferences.RemoveAll(existing => existing.Name.LocalName == reference && existing.Attribute(names.Type)?.Value == type);

        // Header references come before footer references in w:sectPr.
        XElement added = new(names.W + reference, new XAttribute(names.Type, type), new XAttribute(names.Id, id));

        if (header == true)
        {
            int firstFooter = headerReferences.FindIndex(existing => existing.Name.LocalName == "footerReference");

            if (firstFooter < 0)
            {
                headerReferences.Add(added);
            }
            else
            {
                headerReferences.Insert(firstFooter, added);
            }
        }
        else
        {
            headerReferences.Add(added);
        }
    }

    private async Task WriteBlockAsync(XElement block, bool table, CancellationToken cancellationToken)
    {
        ThrowIfCompleted();

        // Two tables next to each other become one in Word; an empty paragraph keeps them apart.
        if (table == true && lastWasTable == true)
        {
            await new XElement(names.Paragraph).WriteToAsync(xml, cancellationToken);
        }

        await block.WriteToAsync(xml, cancellationToken);
        lastWasTable = table;
    }

    private async Task WritePartsAsync(CancellationToken cancellationToken)
    {
        XDocument styles = BuiltInStyles.NewStylesPart(names);
        AddUsedStyles(styles.Root!);
        await WriteXmlAsync("word/styles.xml", WordTypes.StylesContentType, styles, cancellationToken);
        AddRelationship(MainPart, WordTypes.StylesRelationship, "word/styles.xml");

        await WriteXmlAsync("word/settings.xml", WordTypes.SettingsContentType, BuiltInStyles.NewSettingsPart(names), cancellationToken);
        AddRelationship(MainPart, WordTypes.SettingsRelationship, "word/settings.xml");

        if (numberingInstances.Count > 0)
        {
            XDocument numbering = BuiltInStyles.NewNumberingPart(names);
            numbering.Root!.Add(numberingInstances);
            await WriteXmlAsync("word/numbering.xml", WordTypes.NumberingContentType, numbering, cancellationToken);
            AddRelationship(MainPart, WordTypes.NumberingRelationship, "word/numbering.xml");
        }

        await WriteNotesAsync(NoteKind.Footnote, footnotes, cancellationToken);
        await WriteNotesAsync(NoteKind.Endnote, endnotes, cancellationToken);

        foreach ((string part, string contentType, XDocument document) in extraParts)
        {
            await WriteXmlAsync(part, contentType, document, cancellationToken);
        }

        foreach ((string name, string contentType, byte[] data) in media)
        {
            await package.WritePartAsync(name, contentType, data, cancellationToken);
        }

        await WriteCorePropertiesAsync(cancellationToken);
        package.AddRelationship(string.Empty, new OpcRelationship("rId1", OpcNamespaces.OfficeDocumentRelationship, MainPart, false));
        package.AddRelationship(string.Empty, new OpcRelationship("rId2", OpcNamespaces.CorePropertiesRelationship, "docProps/core.xml", false));

        foreach ((string source, List<OpcRelationship> list) in relationships)
        {
            foreach (OpcRelationship relationship in list)
            {
                package.AddRelationship(source, relationship);
            }
        }
    }

    private async Task WriteNotesAsync(NoteKind kind, List<XElement> notes, CancellationToken cancellationToken)
    {
        if (notes.Count == 0)
        {
            return;
        }

        XDocument part = BuiltInStyles.NewNotesPart(names, kind);
        part.Root!.Add(notes);
        string name = "word/footnotes.xml";
        string contentType = WordTypes.FootnotesContentType;
        string relationship = WordTypes.FootnotesRelationship;

        if (kind == NoteKind.Endnote)
        {
            name = "word/endnotes.xml";
            contentType = WordTypes.EndnotesContentType;
            relationship = WordTypes.EndnotesRelationship;
        }

        await WriteXmlAsync(name, contentType, part, cancellationToken);
        AddRelationship(MainPart, relationship, name);
    }

    private async Task WriteCorePropertiesAsync(CancellationToken cancellationToken)
    {
        XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        XNamespace dc = "http://purl.org/dc/elements/1.1/";
        XNamespace dcterms = "http://purl.org/dc/terms/";
        XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";
        string now = DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
        XElement root = new(
            cp + "coreProperties",
            new XAttribute(XNamespace.Xmlns + "cp", cp.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dc", dc.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dcterms", dcterms.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "xsi", xsi.NamespaceName));

        if (options.Title is not null)
        {
            root.Add(new XElement(dc + "title", options.Title));
        }

        if (options.Author is not null)
        {
            root.Add(new XElement(dc + "creator", options.Author));
        }

        root.Add(
            new XElement(dcterms + "created", new XAttribute(xsi + "type", "dcterms:W3CDTF"), now),
            new XElement(dcterms + "modified", new XAttribute(xsi + "type", "dcterms:W3CDTF"), now));

        await WriteXmlAsync("docProps/core.xml", OpcNamespaces.CorePropertiesContentType, Part(root), cancellationToken);
    }

    /// <summary>Adds the styles the content used, and the styles they are based on, to a new styles part.</summary>
    private void AddUsedStyles(XElement root)
    {
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
    }

    private Task WriteXmlAsync(string name, string contentType, XDocument document, CancellationToken cancellationToken)
    {
        return package.WritePartAsync(name, contentType, WordXml.Save(document), cancellationToken);
    }

    private string AddRelationship(string source, string type, string target)
    {
        if (relationships.TryGetValue(source, out List<OpcRelationship>? list) == false)
        {
            list = [];
            relationships[source] = list;
        }

        string id = "rId" + (list.Count + 1).ToString(CultureInfo.InvariantCulture);
        string relative = target;

        if (source.Length > 0)
        {
            relative = OpcPath.RelativeTarget(source, target);
        }

        list.Add(new OpcRelationship(id, type, relative, false));

        return id;
    }

    private string AddExternalRelationship(string source, string type, string url)
    {
        if (relationships.TryGetValue(source, out List<OpcRelationship>? list) == false)
        {
            list = [];
            relationships[source] = list;
        }

        OpcRelationship? existing = list.FirstOrDefault(relationship => relationship.IsExternal == true && relationship.Target == url);

        if (existing is not null)
        {
            return existing.Id;
        }

        string id = "rId" + (list.Count + 1).ToString(CultureInfo.InvariantCulture);
        list.Add(new OpcRelationship(id, type, url, true));

        return id;
    }

    /// <summary>A picture's part, shared by every place the same bytes are used.</summary>
    private string MediaPart(byte[] data, string contentType, string extension)
    {
        string hash = Convert.ToHexString(SHA256.HashData(data));

        if (mediaByHash.TryGetValue(hash, out string? existing) == true)
        {
            return existing;
        }

        string name = "word/media/image" + (media.Count + 1).ToString(CultureInfo.InvariantCulture) + "." + extension;
        media.Add((name, contentType, data));
        mediaByHash[hash] = name;

        return name;
    }

    private string ListNumbering(bool bullet, bool startNew)
    {
        if (bullet == true && bulletList is not null && startNew == false)
        {
            return bulletList;
        }

        if (bullet == false && numberedList is not null && startNew == false)
        {
            return numberedList;
        }

        string id = (numberingInstances.Count + 1).ToString(CultureInfo.InvariantCulture);
        numberingInstances.Add(BuiltInStyles.NumberingInstance(names, id, bullet));

        if (bullet == true)
        {
            bulletList = id;
        }
        else
        {
            numberedList = id;
        }

        return id;
    }

    private void ThrowIfCompleted()
    {
        if (completed == true)
        {
            throw new InvalidOperationException("The document is complete; nothing more can be written to it.");
        }
    }

    /// <summary>The builders' view of the writer, for one part: relationships are the part's own.</summary>
    private sealed class Sink : IContentSink
    {
        private readonly DocxWriter writer;
        private readonly string part;

        public Sink(DocxWriter writer, string part)
        {
            this.writer = writer;
            this.part = part;
        }

        public WordNames Names => writer.names;

        public Length TextWidth => writer.page.TextWidth;

        public string AddHyperlink(string url)
        {
            return writer.AddExternalRelationship(part, OpcNamespaces.HyperlinkRelationship, url);
        }

        public string AddPicture(byte[] data, string contentType, string extension)
        {
            string media = writer.MediaPart(data, contentType, extension);
            string relative = OpcPath.RelativeTarget(part, media);

            if (writer.relationships.TryGetValue(part, out List<OpcRelationship>? list) == true)
            {
                OpcRelationship? existing = list.FirstOrDefault(relationship => relationship.IsExternal == false && relationship.Target == relative);

                if (existing is not null)
                {
                    return existing.Id;
                }
            }

            return writer.AddRelationship(part, OpcNamespaces.ImageRelationship, media);
        }

        public int NextDrawingId()
        {
            writer.drawingId++;

            return writer.drawingId;
        }

        public string AddNote(NoteKind kind, IReadOnlyList<XElement> paragraphs)
        {
            List<XElement> notes = writer.footnotes;
            string element = "footnote";

            if (kind == NoteKind.Endnote)
            {
                notes = writer.endnotes;
                element = "endnote";
            }

            string id = (notes.Count + 1).ToString(CultureInfo.InvariantCulture);
            notes.Add(new XElement(writer.names.W + element, new XAttribute(writer.names.WId, id), paragraphs));

            return id;
        }

        public void UseStyle(string styleId)
        {
            writer.usedStyles.Add(styleId);
        }

        public string ListNumbering(bool bullet, bool startNew)
        {
            return writer.ListNumbering(bullet, startNew);
        }
    }
}
