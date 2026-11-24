namespace Avig.Ord;

/// <summary>
/// Limits on what a file may make the package do, so that a document from outside cannot run
/// the process out of memory or stack. A .docx is a zip file, and a few kilobytes of it can
/// unpack to gigabytes. A file that goes past a limit gives Invalid when it is opened, and its
/// <see cref="OrdError.ErrorKind(Result)"/> is <see cref="OrdErrorKind.TooLarge"/>. Time is limited with the
/// CancellationToken every call takes.
/// </summary>
/// <example>
/// <code>
/// Result&lt;OrdDocument&gt; opened = await OrdDocument.OpenAsync(upload, OrdLimits.Default with { MaxUncompressedBytes = 50_000_000 });
/// </code>
/// </example>
public sealed record OrdLimits
{
    /// <summary>The limits used when none are given: well above any document people write, and well below what takes a server down.</summary>
    public static OrdLimits Default { get; } = new();

    /// <summary>No limits beyond what the format itself has. Only for files you trust.</summary>
    public static OrdLimits Unlimited { get; } = new()
    {
        MaxUncompressedBytes = long.MaxValue,
        MaxParts = int.MaxValue,
        MaxXmlDepth = 10_000,
    };

    /// <summary>
    /// How many bytes the parts of the file may unpack to, together. 512 MiB by default. The
    /// document model holds the XML of the text parts in memory, at several times its size, so
    /// this is lower than in Avig.Kalkyl. Pictures count too.
    /// </summary>
    public long MaxUncompressedBytes { get; init; } = 512L << 20;

    /// <summary>How many parts the package may have. 10 000 by default; Word documents have tens, a few hundred with many pictures.</summary>
    public int MaxParts { get; init; } = 10_000;

    /// <summary>
    /// How deep elements may be nested in the XML. 256 by default: a table inside a table adds
    /// four levels, and Word nests tables far less than 60 deep, while the depth is what decides
    /// how much stack the document model needs.
    /// </summary>
    public int MaxXmlDepth { get; init; } = 256;

    internal Avig.OpenXml.OpcOpenOptions ToOpenOptions()
    {
        return new Avig.OpenXml.OpcOpenOptions
        {
            MaxUncompressedBytes = MaxUncompressedBytes,
            MaxParts = MaxParts,
            MaxXmlDepth = MaxXmlDepth,
            LimitsName = "OrdLimits",
            FormatName = "a Word document",
            CompoundFileHint = "It is either an old .doc document or a password-protected .docx; neither can be read. Save it as an unprotected .docx in Word.",
        };
    }
}
