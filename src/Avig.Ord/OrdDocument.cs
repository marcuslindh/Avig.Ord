using System.Xml;
using System.Xml.Linq;
using Avig.OpenXml;
using Avig.Ord.Internal;

namespace Avig.Ord;

/// <summary>What kind of Word file a document is, from the content type of its main part.</summary>
public enum WordFileKind
{
    /// <summary>A document, .docx.</summary>
    Document,

    /// <summary>A template, .dotx.</summary>
    Template,

    /// <summary>A document with macros, .docm. The macros are kept, never run.</summary>
    MacroEnabledDocument,

    /// <summary>A template with macros, .dotm. The macros are kept, never run.</summary>
    MacroEnabledTemplate,
}

/// <summary>What form a document is saved in.</summary>
public enum WordSaveFormat
{
    /// <summary>As it was opened: a document stays a document, a template a template. A new document is a .docx.</summary>
    Same,

    /// <summary>A .docx document. A macro-enabled file loses its macros.</summary>
    Docx,

    /// <summary>A .dotx template. A macro-enabled file loses its macros.</summary>
    Dotx,

    /// <summary>A .docm document with its macros. A file without macros is a .docm with none.</summary>
    Docm,
}

/// <summary>How a document is saved.</summary>
public sealed record SaveOptions
{
    /// <summary>The options used when none are given.</summary>
    public static SaveOptions Default { get; } = new();

    /// <summary>Document, template or macro-enabled. <see cref="WordSaveFormat.Same"/> by default.</summary>
    public WordSaveFormat Format { get; init; } = WordSaveFormat.Same;

    /// <summary>Save as Flat OPC (the whole document as one XML file, "Word XML Document") instead of a zip.</summary>
    public bool FlatOpc { get; init; }
}

/// <summary>
/// A Word document in memory: open one, read and change it, and save it. Everything the package
/// does not understand — SmartArt, charts, equations, settings, custom XML, the parts of other
/// programs — is kept as it was and saved again unchanged.
/// </summary>
/// <example>
/// <code>
/// Result&lt;OrdDocument&gt; opened = await OrdDocument.OpenAsync("report.docx");
///
/// if (opened.IsSuccess(out OrdDocument? document) == true)
/// {
///     foreach (Paragraph paragraph in document.Paragraphs.Where(paragraph => paragraph.IsHeading == true))
///     {
///         Console.WriteLine(paragraph.Text);
///     }
/// }
/// </code>
/// </example>
public sealed partial class OrdDocument
{
    private readonly Dictionary<string, PartData> parts;
    private readonly List<string> partOrder;
    private readonly Dictionary<string, List<OpcRelationship>> relationships;
    private readonly Dictionary<string, DocumentContext> contexts = new(StringComparer.OrdinalIgnoreCase);
    private StyleSheet? styles;
    private NumberingDefinitions? numbering;

    private OrdDocument(ReadOptions options, string mainPart, Dictionary<string, PartData> parts, List<string> partOrder, Dictionary<string, List<OpcRelationship>> relationships)
    {
        Options = options;
        MainPartName = mainPart;
        this.parts = parts;
        this.partOrder = partOrder;
        this.relationships = relationships;
    }

    /// <summary>How the document was read: the revision view and the limits.</summary>
    public ReadOptions Options { get; }

    /// <summary>The main part, usually word/document.xml.</summary>
    public string MainPartName { get; }

    /// <summary>Document, template or macro-enabled.</summary>
    public WordFileKind Kind => KindOf(parts[MainPartName].ContentType);

    /// <summary>True when the file holds a VBA project. The macros are never run; <see cref="RemoveMacros"/> takes them out.</summary>
    public bool HasMacros => parts.Values.Any(part => string.Equals(part.ContentType, OpcNamespaces.VbaProjectContentType, StringComparison.OrdinalIgnoreCase) == true)
        || parts.Keys.Any(name => name.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase) == true);

    /// <summary>The styles.</summary>
    public StyleSheet Styles => styles ??= StyleSheet.Read(FindPartXml(MainPartName, "styles")?.Root);

    /// <summary>The paragraphs and tables of the body, in order.</summary>
    public IReadOnlyList<Block> Blocks => Block.Read(MainContext, Body);

    /// <summary>Every paragraph of the body in order, also those in tables (and tables in tables).</summary>
    public IReadOnlyList<Paragraph> Paragraphs => [.. AllParagraphs(Blocks)];

    /// <summary>The tables directly in the body.</summary>
    public IReadOnlyList<Table> Tables => [.. Blocks.OfType<Table>()];

    /// <summary>The document's text as the revision view has it: paragraphs on lines of their own, table cells separated by tabs.</summary>
    public string Text => TextRenderer.Render(Blocks, TextOptions.Default);

    internal NumberingDefinitions Numbering => numbering ??= NumberingDefinitions.Read(FindPartXml(MainPartName, "numbering")?.Root);

    internal DocumentContext MainContext => ContextFor(MainPartName);

    internal XElement Body
    {
        get
        {
            XDocument document = Xml(MainPartName);
            WordNames names = WordNames.For(document.Root);
            XElement root = document.Root ?? throw new InvalidOperationException("The main part has no root element.");
            XElement? body = root.Element(names.Body);

            if (body is null)
            {
                body = new XElement(names.Body);
                root.Add(body);
            }

            return body;
        }
    }

    /// <summary>Opens a file.</summary>
    /// <param name="path">The file.</param>
    /// <param name="options">The revision view and limits; <see cref="ReadOptions.Default"/> when null.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>
    /// The document; NotFound when the file does not exist, Forbidden when it may not be read,
    /// Conflict when another program holds it locked, and Invalid when it is not a Word
    /// document, is damaged, or goes past a limit (see <see cref="OrdError.ErrorKind(Result)"/>).
    /// </returns>
    public static async Task<Result<OrdDocument>> OpenAsync(string path, ReadOptions? options = null, CancellationToken cancellationToken = default)
    {
        options ??= ReadOptions.Default;
        Result<OpcPackage> opened = await OpcPackage.OpenAsync(path, options.Limits.ToOpenOptions(), cancellationToken);

        return await LoadAsync(opened, options, cancellationToken);
    }

    /// <summary>Reads a document from a stream, which is read to the end and not kept.</summary>
    /// <param name="input">The stream.</param>
    /// <param name="options">The revision view and limits; <see cref="ReadOptions.Default"/> when null.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>The document, or Invalid as for <see cref="OpenAsync(string, ReadOptions?, CancellationToken)"/>.</returns>
    public static async Task<Result<OrdDocument>> OpenAsync(Stream input, ReadOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        options ??= ReadOptions.Default;
        Result<OpcPackage> opened = await OpcPackage.OpenAsync(input, options.Limits.ToOpenOptions(), true, cancellationToken);

        return await LoadAsync(opened, options, cancellationToken);
    }

    /// <summary>Saves the document to a file, replacing it if it exists.</summary>
    /// <param name="path">The file.</param>
    /// <param name="options">The format; <see cref="SaveOptions.Default"/> when null.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>Success; Forbidden when the file may not be written, Conflict when another program holds it.</returns>
    public async Task<Result> SaveAsync(string path, SaveOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        FileStream file;

        try
        {
            file = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 65_536, FileOptions.Asynchronous);
        }
        catch (DirectoryNotFoundException)
        {
            return Result.NotFound($"The folder for '{path}' does not exist.");
        }
        catch (UnauthorizedAccessException)
        {
            return Result.Forbidden($"'{path}' may not be written.");
        }
        catch (IOException exception)
        {
            return Result.Conflict($"'{path}' could not be written: {exception.Message}");
        }

        await using (file)
        {
            await SaveAsync(file, options, cancellationToken);
        }

        return Result.Success();
    }

    /// <summary>Saves the document to a stream, which is left open.</summary>
    /// <param name="output">The stream.</param>
    /// <param name="options">The format; <see cref="SaveOptions.Default"/> when null.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>A task that completes when the document is written.</returns>
    public async Task SaveAsync(Stream output, SaveOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(output);
        options ??= SaveOptions.Default;
        BeforeSave();
        OpcPackageKind kind = OpcPackageKind.Zip;

        if (options.FlatOpc == true)
        {
            kind = OpcPackageKind.FlatOpc;
        }

        string mainType = SavedMainContentType(options.Format);
        bool dropMacros = mainType is WordTypes.DocumentContentType or WordTypes.TemplateContentType;
        HashSet<string> dropped = new(StringComparer.OrdinalIgnoreCase);

        if (dropMacros == true)
        {
            dropped = MacroParts();
        }

        await using OpcPackageWriter writer = await OpcPackageWriter.CreateAsync(output, kind, true, cancellationToken);
        writer.FlatOpcProgId = "Word.Document";
        writer.MaxXmlDepth = Options.Limits.MaxXmlDepth;

        foreach (string name in partOrder)
        {
            if (dropped.Contains(name) == true || parts.TryGetValue(name, out PartData? part) == false)
            {
                continue;
            }

            string contentType = part.ContentType;

            if (string.Equals(name, MainPartName, StringComparison.OrdinalIgnoreCase) == true)
            {
                contentType = mainType;
            }

            byte[] data = part.Data;

            if (part.Xml is not null && part.Dirty == true)
            {
                data = WordXml.Save(part.Xml);
            }

            await writer.WritePartAsync(name, contentType, data, cancellationToken);
        }

        foreach ((string source, List<OpcRelationship> list) in relationships)
        {
            if (source.Length > 0 && (parts.ContainsKey(source) == false || dropped.Contains(source) == true))
            {
                continue;
            }

            foreach (OpcRelationship relationship in list)
            {
                if (relationship.IsExternal == false && dropped.Contains(OpcPath.ResolveTarget(source, relationship.Target)) == true)
                {
                    continue;
                }

                writer.AddRelationship(source, relationship);
            }
        }

        await writer.CompleteAsync(cancellationToken);
    }

    /// <summary>Takes out the VBA project and everything that belongs to it. The file is then saved as a .docx (or .dotx) unless another format is asked for.</summary>
    public void RemoveMacros()
    {
        foreach (string name in MacroParts())
        {
            parts.Remove(name);
            relationships.Remove(name);
        }

        PartData main = parts[MainPartName];

        if (main.ContentType == WordTypes.MacroDocumentContentType)
        {
            main.ContentType = WordTypes.DocumentContentType;
        }
        else if (main.ContentType == WordTypes.MacroTemplateContentType)
        {
            main.ContentType = WordTypes.TemplateContentType;
        }
    }

    /// <summary>Work done before the parts are written: styles and lists the document now uses are added to their parts.</summary>
    partial void BeforeSave();

    /// <summary>Every paragraph among the blocks, in order, also those in tables.</summary>
    internal static IEnumerable<Paragraph> AllParagraphs(IEnumerable<Block> blocks)
    {
        foreach (Block block in blocks)
        {
            if (block is Paragraph paragraph)
            {
                yield return paragraph;

                continue;
            }

            if (block is not Table table)
            {
                continue;
            }

            foreach (TableRow row in table.Rows)
            {
                foreach (TableCell cell in row.Cells)
                {
                    foreach (Paragraph inner in AllParagraphs(cell.Blocks))
                    {
                        yield return inner;
                    }
                }
            }
        }
    }

    /// <summary>Counts the list labels of a part from its start, if anything changed since they were last counted.</summary>
    internal void EnsureListLabels(DocumentContext context)
    {
        if (context.LabelledVersion == context.Version)
        {
            return;
        }

        context.ResetListCounts();
        XElement? root = Xml(context.PartName).Root;

        if (root is not null)
        {
            foreach (Paragraph paragraph in AllParagraphs(Block.Read(context, ContainerOf(context, root))))
            {
                context.CountListItem(paragraph.Element, paragraph.StyleId, paragraph.Properties);
            }
        }

        context.LabelledVersion = context.Version;
    }

    /// <summary>The context for the content of a part: the body, a header, the notes …</summary>
    internal DocumentContext ContextFor(string partName)
    {
        if (contexts.TryGetValue(partName, out DocumentContext? context) == true)
        {
            return context;
        }

        XDocument document = Xml(partName);
        context = new DocumentContext(WordNames.For(document.Root), Styles, Numbering, partName, RelationshipsOf(partName), Options.Revisions, this, ReadPartAsync);
        contexts[partName] = context;

        return context;
    }

    /// <summary>A part's XML, loaded the first time it is asked for and marked as changed, since whoever asks may change it.</summary>
    internal XDocument Xml(string partName)
    {
        PartData part = parts[partName];

        if (part.Xml is null)
        {
            part.Xml = WordXml.Load(part.Data, Options.Limits.MaxXmlDepth);
        }

        part.Dirty = true;

        return part.Xml;
    }

    internal bool HasPart(string partName)
    {
        return parts.ContainsKey(partName);
    }

    internal IReadOnlyList<OpcRelationship> RelationshipsOf(string partName)
    {
        if (relationships.TryGetValue(partName, out List<OpcRelationship>? list) == false)
        {
            list = [];
            relationships[partName] = list;
        }

        return list;
    }

    /// <summary>The XML of the part the first relationship of a kind points to from <paramref name="source"/>; null when there is none.</summary>
    internal XDocument? FindPartXml(string source, string relationshipKind)
    {
        XDocument? found = null;

        foreach (OpcRelationship relationship in RelationshipsOf(source))
        {
            if (relationship.IsExternal == false && OpcPath.IsRelationship(relationship.Type, relationshipKind) == true)
            {
                string target = OpcPath.ResolveTarget(source, relationship.Target);

                if (parts.ContainsKey(target) == true)
                {
                    found = Xml(target);

                    break;
                }
            }
        }

        return found;
    }

    private static WordFileKind KindOf(string contentType)
    {
        return contentType switch
        {
            WordTypes.TemplateContentType => WordFileKind.Template,
            WordTypes.MacroDocumentContentType => WordFileKind.MacroEnabledDocument,
            WordTypes.MacroTemplateContentType => WordFileKind.MacroEnabledTemplate,
            _ => WordFileKind.Document,
        };
    }

    private static async Task<Result<OrdDocument>> LoadAsync(Result<OpcPackage> opened, ReadOptions options, CancellationToken cancellationToken)
    {
        if (opened.IsFailure(out string? error, out ResultStatus status, out OpcPackage? package) == true)
        {
            return Result<OrdDocument>.Failure(error, status);
        }

        await using (package)
        {
            try
            {
                return await ReadPackageAsync(package, options, cancellationToken);
            }
            catch (XmlException exception)
            {
                return Result<OrdDocument>.Invalid($"{OpcErrors.DamagedPrefix}: a part has broken XML ({exception.Message}).");
            }
            catch (InvalidDataException exception)
            {
                return Result<OrdDocument>.Invalid(OpcPackage.Describe(exception));
            }
        }
    }

    private static async Task<Result<OrdDocument>> ReadPackageAsync(OpcPackage package, ReadOptions options, CancellationToken cancellationToken)
    {
        Result<string> main = await package.FindMainPartAsync(cancellationToken);

        if (main.IsFailure(out _, out _, out string? mainPart) == true || package.Contains(mainPart) == false)
        {
            return Result<OrdDocument>.Invalid($"{OrdError.NotWordPrefix}: it is an Open XML package, but has no main document part.");
        }

        package.TryGetContentType(mainPart, out string? mainType);

        if (WordTypes.IsMainContentType(mainType) == false)
        {
            return Result<OrdDocument>.Invalid($"{OrdError.NotWordPrefix}: its main part is '{mainType ?? "of no stated type"}', not a Word document. It may be a spreadsheet or a presentation.");
        }

        Dictionary<string, PartData> parts = new(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, List<OpcRelationship>> relationships = new(StringComparer.OrdinalIgnoreCase);
        List<string> order = [];

        foreach (string name in package.PartNames)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (string.Equals(name, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase) == true)
            {
                continue;
            }

            if (name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) == true && name.Contains("_rels/", StringComparison.OrdinalIgnoreCase) == true)
            {
                string source = SourceOfRelationships(name);
                List<OpcRelationship> list = [.. await package.GetRelationshipsAsync(source, cancellationToken)];
                HashSet<string> ids = new(StringComparer.Ordinal);

                // Two relationships with one id cannot be told apart, and cannot be saved.
                foreach (OpcRelationship relationship in list)
                {
                    if (ids.Add(relationship.Id) == false)
                    {
                        return Result<OrdDocument>.Invalid($"{OpcErrors.DamagedPrefix}: the relationships of '{source}' have the id '{relationship.Id}' twice.");
                    }
                }

                relationships[source] = list;

                continue;
            }

            Result<byte[]> data = await package.ReadPartAsync(name, cancellationToken);

            if (data.IsSuccess(out byte[]? bytes) == true)
            {
                if (package.TryGetContentType(name, out string? type) == false)
                {
                    type = "application/octet-stream";
                }

                parts[name] = new PartData(type, bytes);
                order.Add(name);
            }
        }

        // Every XML part is read through once now, with the file's limits, so that broken XML
        // anywhere is found when the file is opened: the model loads the other parts only when
        // they are asked for (the styles, a header), and a Flat OPC save writes every XML part
        // as XML. Without this, OpenAsync succeeded and Headers or SaveAsync threw XmlException.
        foreach (PartData part in parts.Values)
        {
            CheckXml(part, options.Limits.MaxXmlDepth, cancellationToken);
        }

        OrdDocument document = new(options, mainPart, parts, order, relationships);

        // Load the main part now, so that broken XML in it is found when the file is opened.
        _ = document.Body;
        parts[mainPart].Dirty = false;

        return document;
    }

    /// <summary>
    /// Reads an XML part through without keeping it, and throws <see cref="XmlException"/> (or
    /// <see cref="InvalidDataException"/> past the depth limit) when it is broken. An empty part
    /// is broken only when it is one of the parts the model reads as XML; any other empty part
    /// is saved as it is.
    /// </summary>
    private static void CheckXml(PartData part, int maxDepth, CancellationToken cancellationToken)
    {
        string type = part.ContentType;
        bool xml = type.EndsWith("+xml", StringComparison.OrdinalIgnoreCase) == true || type.EndsWith("/xml", StringComparison.OrdinalIgnoreCase) == true;
        bool word = type.StartsWith("application/vnd.openxmlformats-officedocument.wordprocessingml.", StringComparison.OrdinalIgnoreCase) == true
            || WordTypes.IsMainContentType(type) == true
            || string.Equals(type, OpcNamespaces.CorePropertiesContentType, StringComparison.OrdinalIgnoreCase) == true;

        if (xml == false || (part.Data.Length == 0 && word == false))
        {
            return;
        }

        using MemoryStream stream = new(part.Data, false);
        using XmlReader reader = OpcXml.CreateReader(stream, maxDepth);

        while (reader.Read() == true)
        {
            cancellationToken.ThrowIfCancellationRequested();
        }
    }

    /// <summary>The part whose relationships a .rels part holds: word/_rels/document.xml.rels is word/document.xml's, _rels/.rels the package's.</summary>
    private static string SourceOfRelationships(string relsPart)
    {
        int marker = relsPart.LastIndexOf("_rels/", StringComparison.OrdinalIgnoreCase);
        string folder = relsPart[..marker];
        string file = relsPart[(marker + "_rels/".Length)..^".rels".Length];

        return folder + file;
    }

    /// <summary>The container of blocks in a part: the body of the document, the root of a header, footer or note part's notes.</summary>
    private static XElement ContainerOf(DocumentContext context, XElement root)
    {
        return root.Element(context.Names.Body) ?? root;
    }

    private string SavedMainContentType(WordSaveFormat format)
    {
        string current = parts[MainPartName].ContentType;

        return format switch
        {
            WordSaveFormat.Docx => WordTypes.DocumentContentType,
            WordSaveFormat.Dotx => WordTypes.TemplateContentType,
            WordSaveFormat.Docm => WordTypes.MacroDocumentContentType,
            _ => current,
        };
    }

    /// <summary>The VBA project, its signature and its data, found by content type and by what the main part's vbaProject relationship points to.</summary>
    private HashSet<string> MacroParts()
    {
        HashSet<string> macros = new(StringComparer.OrdinalIgnoreCase);

        foreach ((string name, PartData part) in parts)
        {
            if (part.ContentType.StartsWith("application/vnd.ms-office.vba", StringComparison.OrdinalIgnoreCase) == true
                || name.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase) == true
                || name.EndsWith("vbaData.xml", StringComparison.OrdinalIgnoreCase) == true)
            {
                macros.Add(name);
            }
        }

        return macros;
    }

    private async Task<Result<byte[]>> ReadPartAsync(string partName, CancellationToken cancellationToken)
    {
        await Task.CompletedTask;
        cancellationToken.ThrowIfCancellationRequested();

        if (parts.TryGetValue(partName, out PartData? part) == false)
        {
            return Result<byte[]>.NotFound($"The document has no part '{partName}'.");
        }

        return part.Data.ToArray();
    }

    /// <summary>A part of the package as it is held: its bytes as read, and its XML once loaded.</summary>
    internal sealed class PartData(string contentType, byte[] data)
    {
        public string ContentType { get; set; } = contentType;

        public byte[] Data { get; set; } = data;

        public XDocument? Xml { get; set; }

        /// <summary>True once the XML has been handed out and may have been changed; the part is then saved from its XML.</summary>
        public bool Dirty { get; set; }
    }
}
