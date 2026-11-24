using Avig.OpenXml;

namespace Avig.Ord;

/// <summary>
/// What kind of failure a result or exception from the package is, for an application to show
/// to its users. The error message is meant for the log: it can name parts of the file and
/// quote text from it. The kind says no more than what went wrong in general, so it is safe to
/// show, and an application can word it in its own language.
/// </summary>
public enum OrdErrorKind
{
    /// <summary>Not a failure the package recognizes: a success, or an error that did not come from opening or reading a file.</summary>
    Unknown,

    /// <summary>The file or its folder does not exist.</summary>
    NotFound,

    /// <summary>The file is not a Word document: an old .doc, a password-protected .docx, another kind of file, or not a zip or Flat OPC file at all.</summary>
    NotWord,

    /// <summary>The file is a Word document, or claims to be one, but is damaged or holds something no word processor writes.</summary>
    Damaged,

    /// <summary>The file goes past one of the <see cref="OrdLimits"/>. A trusted file can be opened with larger limits.</summary>
    TooLarge,

    /// <summary>The file exists but could not be opened or written: another program has it open, or the process may not use it.</summary>
    Inaccessible,
}

/// <summary>Finds the <see cref="OrdErrorKind"/> of a failed result or of an exception thrown while a file is read.</summary>
/// <example>
/// <code>
/// Result&lt;OrdDocument&gt; opened = await OrdDocument.OpenAsync(upload);
///
/// if (opened.IsFailure(out string? error, out _) == true)
/// {
///     logger.LogWarning("Upload rejected: {Error}", error);
///
///     return opened.ErrorKind() switch
///     {
///         OrdErrorKind.TooLarge => "The document is too large.",
///         OrdErrorKind.NotWord => "Upload a .docx file.",
///         _ => "The document could not be read.",
///     };
/// }
/// </code>
/// </example>
public static class OrdError
{
    /// <summary>How every message about a file that is not a Word document starts.</summary>
    internal const string NotWordPrefix = "The file is not a Word document";

    /// <summary>The kind of a failed result; <see cref="OrdErrorKind.Unknown"/> for a success.</summary>
    public static OrdErrorKind ErrorKind(this Result result)
    {
        return Classify(result.Status, result.Error);
    }

    /// <summary>The kind of a failed result; <see cref="OrdErrorKind.Unknown"/> for a success.</summary>
    public static OrdErrorKind ErrorKind<T>(this Result<T> result)
    {
        return Classify(result.Status, result.Error);
    }

    /// <summary>
    /// The kind of an exception thrown while a file is read, as <see cref="DocxReader"/> does with
    /// <see cref="InvalidDataException"/> for a part that turns out to be damaged. Any other
    /// exception, a cancellation among them, is <see cref="OrdErrorKind.Unknown"/>.
    /// </summary>
    public static OrdErrorKind ErrorKind(this Exception exception)
    {
        ArgumentNullException.ThrowIfNull(exception);

        if (exception is FileNotFoundException or DirectoryNotFoundException)
        {
            return OrdErrorKind.NotFound;
        }

        if (exception is InvalidDataException or System.Xml.XmlException)
        {
            OrdErrorKind kind = OfMessage(exception.Message);

            if (kind == OrdErrorKind.Unknown)
            {
                return OrdErrorKind.Damaged;
            }

            return kind;
        }

        return OrdErrorKind.Unknown;
    }

    private static OrdErrorKind Classify(ResultStatus status, string? error)
    {
        if (status == ResultStatus.Ok || error is null)
        {
            return OrdErrorKind.Unknown;
        }

        if (status == ResultStatus.NotFound)
        {
            return OrdErrorKind.NotFound;
        }

        if (status is ResultStatus.Conflict or ResultStatus.Forbidden)
        {
            return OrdErrorKind.Inaccessible;
        }

        return OfMessage(error);
    }

    private static OrdErrorKind OfMessage(string message)
    {
        if (message.Contains(OpcErrors.LimitMarker + "OrdLimits.", StringComparison.Ordinal) == true)
        {
            return OrdErrorKind.TooLarge;
        }

        if (message.StartsWith(OpcErrors.DamagedPrefix, StringComparison.Ordinal) == true)
        {
            return OrdErrorKind.Damaged;
        }

        if (message.StartsWith(NotWordPrefix, StringComparison.Ordinal) == true)
        {
            return OrdErrorKind.NotWord;
        }

        return OrdErrorKind.Unknown;
    }
}
