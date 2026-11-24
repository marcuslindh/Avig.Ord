using System.Runtime.CompilerServices;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Avig.OpenXml;
using Avig.Ord.Internal;

namespace Avig.Ord;

/// <summary>
/// Reads a document's body block by block without loading it: one paragraph or one table at a
/// time, as read-only <see cref="Paragraph"/> and <see cref="Table"/> objects of the same types
/// <see cref="OrdDocument"/> has. Memory stays at what the largest block takes. For a table so
/// large that even one block is too much, <see cref="ReadEventsAsync"/> reads element by element.
/// </summary>
/// <example>
/// <code>
/// Result&lt;DocxReader&gt; opened = await DocxReader.OpenAsync("big.docx");
///
/// if (opened.IsSuccess(out DocxReader? reader) == true)
/// {
///     await using (reader)
///     {
///         await foreach (Block block in reader.ReadBlocksAsync())
///         {
///             Console.WriteLine(block.Text);
///         }
///     }
/// }
/// </code>
/// </example>
public sealed class DocxReader : IAsyncDisposable
{
    private readonly OpcPackage package;
    private readonly string mainPart;
    private readonly DocumentContext context;
    private bool started;

    private DocxReader(OpcPackage package, string mainPart, DocumentContext context, ReadOptions options)
    {
        this.package = package;
        this.mainPart = mainPart;
        this.context = context;
        Options = options;
    }

    /// <summary>How the document is read.</summary>
    public ReadOptions Options { get; }

    /// <summary>The document's styles.</summary>
    public StyleSheet Styles => context.Styles;

    /// <summary>True when the file holds a VBA project. The macros are never run.</summary>
    public bool HasMacros => package.HasMacros;

    /// <summary>Opens a file. The styles and list definitions are read now; the body as it is asked for.</summary>
    /// <param name="path">The file.</param>
    /// <param name="options">The revision view and limits; <see cref="ReadOptions.Default"/> when null.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>The reader, or a failure as for <see cref="OrdDocument.OpenAsync(string, ReadOptions?, CancellationToken)"/>.</returns>
    public static async Task<Result<DocxReader>> OpenAsync(string path, ReadOptions? options = null, CancellationToken cancellationToken = default)
    {
        options ??= ReadOptions.Default;
        Result<OpcPackage> opened = await OpcPackage.OpenAsync(path, options.Limits.ToOpenOptions(), cancellationToken);

        return await StartAsync(opened, options, cancellationToken);
    }

    /// <summary>Reads from a stream, which the reader disposes unless <paramref name="leaveOpen"/>.</summary>
    /// <param name="input">The stream.</param>
    /// <param name="options">The revision view and limits; <see cref="ReadOptions.Default"/> when null.</param>
    /// <param name="leaveOpen">Leave the stream open when the reader is disposed.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>The reader, or a failure.</returns>
    public static async Task<Result<DocxReader>> OpenAsync(Stream input, ReadOptions? options = null, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        options ??= ReadOptions.Default;
        Result<OpcPackage> opened = await OpcPackage.OpenAsync(input, options.Limits.ToOpenOptions(), leaveOpen, cancellationToken);

        return await StartAsync(opened, options, cancellationToken);
    }

    /// <summary>
    /// The body's paragraphs and tables, in order, read-only. A block content control's
    /// paragraphs come as the body's own. List labels are counted as the reader goes. A part
    /// that turns out to be damaged throws an <see cref="InvalidDataException"/> or an
    /// <see cref="XmlException"/>, whose <see cref="OrdError.ErrorKind(Exception)"/> says so.
    /// </summary>
    /// <exception cref="InvalidOperationException">The body is already being read: it can be read once.</exception>
    public async IAsyncEnumerable<Block> ReadBlocksAsync([EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        await using Stream part = await OpenBodyAsync(cancellationToken);
        using XmlReader reader = package.CreateXmlReader(part);

        if (await MoveToBodyAsync(reader) == false)
        {
            yield break;
        }

        int bodyDepth = reader.Depth;
        await reader.ReadAsync();

        while (reader.EOF == false && OpcXml.IsEnd(reader, bodyDepth) == false)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (reader.NodeType != XmlNodeType.Element)
            {
                await reader.ReadAsync();

                continue;
            }

            if (reader.LocalName == "sectPr")
            {
                await reader.SkipAsync();

                continue;
            }

            XNode node = await XNode.ReadFromAsync(reader, cancellationToken);

            if (node is not XElement element)
            {
                continue;
            }

            XElement wrapper = new(context.Names.Body, element);

            foreach (Block block in Block.Read(context, wrapper))
            {
                foreach (Paragraph paragraph in OrdDocument.AllParagraphs([block]))
                {
                    context.CountListItem(paragraph.Element, paragraph.StyleId, paragraph.Properties);
                }

                yield return block;
            }
        }
    }

    /// <summary>The whole body as plain text, read block by block.</summary>
    public async Task<string> ReadPlainTextAsync(CancellationToken cancellationToken = default)
    {
        StringBuilder text = new();

        await foreach (Block block in ReadBlocksAsync(cancellationToken))
        {
            TextRenderer.Write(text, [block], TextOptions.Default, 0);
        }

        // As TextRenderer.Render: only the last paragraph's own line end goes, so that an empty
        // last paragraph, or a line break at the end, reads as in OrdDocument.ToPlainText.
        if (text.Length > 0 && text[^1] == '\n')
        {
            text.Length--;
        }

        return text.ToString();
    }

    /// <summary>The whole body as Markdown, written to <paramref name="output"/> block by block.</summary>
    public async Task WriteMarkdownAsync(TextWriter output, MarkdownOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(output);
        options ??= MarkdownOptions.Default;

        // One renderer for the whole body, so that it knows what came before as the document
        // model's ToMarkdown does: an empty paragraph between two list items used to end the list.
        MarkdownRenderer renderer = MarkdownRenderer.ForStream(options);

        await foreach (Block block in ReadBlocksAsync(cancellationToken))
        {
            await output.WriteAsync(renderer.Next(block).AsMemory(), cancellationToken);
        }

        await output.WriteAsync(MarkdownRenderer.Finish().AsMemory(), cancellationToken);
    }

    /// <summary>
    /// The body element by element, for a document whose tables are too large to read one at a
    /// time: where paragraphs, tables, rows and cells start and end, and the text in between, in
    /// the reader's revision view. Field instructions are left out and their results kept.
    /// </summary>
    /// <exception cref="InvalidOperationException">The body is already being read: it can be read once.</exception>
    public async IAsyncEnumerable<DocumentEvent> ReadEventsAsync([EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        await using Stream part = await OpenBodyAsync(cancellationToken);
        using XmlReader reader = package.CreateXmlReader(part);

        if (await MoveToBodyAsync(reader) == false)
        {
            yield break;
        }

        await foreach (DocumentEvent item in EventReader.ReadAsync(reader, context.View, cancellationToken))
        {
            yield return item;
        }
    }

    /// <summary>Closes the file.</summary>
    /// <returns>A task that completes when the file is closed.</returns>
    public ValueTask DisposeAsync()
    {
        return package.DisposeAsync();
    }

    private static async Task<Result<DocxReader>> StartAsync(Result<OpcPackage> opened, ReadOptions options, CancellationToken cancellationToken)
    {
        if (opened.IsFailure(out string? error, out ResultStatus status, out OpcPackage? package) == true)
        {
            return Result<DocxReader>.Failure(error, status);
        }

        try
        {
            Result<DocxReader> reader = await StartAsync(package, options, cancellationToken);

            if (reader.IsFailure(out _, out _) == true)
            {
                await package.DisposeAsync();
            }

            return reader;
        }
        catch (XmlException exception)
        {
            await package.DisposeAsync();

            return Result<DocxReader>.Invalid($"{OpcErrors.DamagedPrefix}: a part has broken XML ({exception.Message}).");
        }
        catch (InvalidDataException exception)
        {
            await package.DisposeAsync();

            return Result<DocxReader>.Invalid(OpcPackage.Describe(exception));
        }
    }

    private static async Task<Result<DocxReader>> StartAsync(OpcPackage package, ReadOptions options, CancellationToken cancellationToken)
    {
        Result<string> main = await package.FindMainPartAsync(cancellationToken);

        if (main.IsFailure(out _, out _, out string? mainPart) == true || package.Contains(mainPart) == false)
        {
            return Result<DocxReader>.Invalid($"{OrdError.NotWordPrefix}: it is an Open XML package, but has no main document part.");
        }

        package.TryGetContentType(mainPart, out string? mainType);

        if (WordTypes.IsMainContentType(mainType) == false)
        {
            return Result<DocxReader>.Invalid($"{OrdError.NotWordPrefix}: its main part is '{mainType ?? "of no stated type"}', not a Word document.");
        }

        XElement? stylesRoot = await RelatedRootAsync(package, mainPart, "styles", options, cancellationToken);
        XElement? numberingRoot = await RelatedRootAsync(package, mainPart, "numbering", options, cancellationToken);
        StyleSheet styles = StyleSheet.Read(stylesRoot);
        NumberingDefinitions numbering = NumberingDefinitions.Read(numberingRoot);
        IReadOnlyList<OpcRelationship> relationships = await package.GetRelationshipsAsync(mainPart, cancellationToken);
        WordNames names = await NamesOfAsync(package, mainPart, options, cancellationToken);

        DocumentContext context = new(names, styles, numbering, mainPart, relationships, options.Revisions, null, package.ReadPartAsync);

        return new DocxReader(package, mainPart, context, options);
    }

    private static async Task<XElement?> RelatedRootAsync(OpcPackage package, string source, string kind, ReadOptions options, CancellationToken cancellationToken)
    {
        XElement? root = null;
        Result<string> part = await package.FindRelatedPartAsync(source, kind, cancellationToken);

        if (part.IsSuccess(out string? name) == true && (await package.ReadPartAsync(name, cancellationToken)).IsSuccess(out byte[]? data) == true)
        {
            root = WordXml.Load(data, options.Limits.MaxXmlDepth).Root;
        }

        return root;
    }

    /// <summary>The namespace of the main part, read from its root element only.</summary>
    private static async Task<WordNames> NamesOfAsync(OpcPackage package, string mainPart, ReadOptions options, CancellationToken cancellationToken)
    {
        WordNames names = WordNames.Transitional;
        Result<Stream> opened = await package.OpenPartAsync(mainPart, cancellationToken);

        if (opened.IsSuccess(out Stream? stream) == true)
        {
            await using (stream)
            {
                using XmlReader reader = OpcXml.CreateReader(stream, options.Limits.MaxXmlDepth);

                if (await reader.MoveToContentAsync() == XmlNodeType.Element && reader.NamespaceURI == WordNames.StrictMain)
                {
                    names = WordNames.Strict;
                }
            }
        }

        return names;
    }

    private static async Task<bool> MoveToBodyAsync(XmlReader reader)
    {
        while (await reader.ReadAsync() == true)
        {
            if (OpcXml.IsElement(reader, "body") == true)
            {
                return reader.IsEmptyElement == false;
            }
        }

        return false;
    }

    private async Task<Stream> OpenBodyAsync(CancellationToken cancellationToken)
    {
        if (started == true)
        {
            throw new InvalidOperationException("The body is already read or being read. DocxReader reads it once; open the file again to read it again.");
        }

        started = true;
        Result<Stream> opened = await package.OpenPartAsync(mainPart, cancellationToken);

        if (opened.IsFailure(out string? error, out _, out Stream? stream) == true)
        {
            throw new InvalidDataException($"{OpcErrors.DamagedPrefix}: {error}");
        }

        return stream;
    }
}
