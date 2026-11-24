using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace Avig.Ord.Fuzz;

/// <summary>Microsoft's schema validator, as <c>TestFiles.Validate</c> in the tests runs it.</summary>
internal static class Validator
{
    /// <summary>The errors the validator finds, one per line; empty when the file is valid.</summary>
    public static string Validate(byte[] file)
    {
        using MemoryStream stream = new(file, false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, false);
        OpenXmlValidator validator = new(FileFormatVersions.Microsoft365);
        IEnumerable<ValidationErrorInfo> errors = validator.Validate(document);

        return string.Join('\n', errors.Take(10).Select(error => $"{error.Part?.Uri} {error.Path?.XPath}: {error.Description}"));
    }

    /// <summary>True when the Open XML SDK can open the file and its validator finds nothing. A file the SDK cannot open is not valid.</summary>
    public static bool IsValid(byte[] file)
    {
        try
        {
            return Validate(file).Length == 0;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException)
        {
            // The SDK throws for files it cannot read (broken XML, a missing main part): those are not valid input.
            return false;
        }
    }
}
