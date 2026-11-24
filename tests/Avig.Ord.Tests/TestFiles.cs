using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using WordDocument = DocumentFormat.OpenXml.Wordprocessing.Document;

namespace Avig.Ord.Tests;

/// <summary>Documents made by Microsoft's Open XML SDK, as an independent source of files, and its schema validator.</summary>
internal static class TestFiles
{
    public const string Wml = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>A .docx whose body is <paramref name="bodyXml"/> (w: prefix declared), with optional styles and numbering parts.</summary>
    public static byte[] Docx(string bodyXml, string? stylesXml = null, string? numberingXml = null, Action<MainDocumentPart>? extra = null)
    {
        using MemoryStream file = new();

        using (WordprocessingDocument document = WordprocessingDocument.Create(file, WordprocessingDocumentType.Document))
        {
            MainDocumentPart main = document.AddMainDocumentPart();
            string xml = $"<w:document xmlns:w=\"{Wml}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><w:body>{bodyXml}</w:body></w:document>";

            using (Stream stream = main.GetStream(FileMode.Create))
            using (StreamWriter writer = new(stream))
            {
                writer.Write(xml);
            }

            if (stylesXml is not null)
            {
                StyleDefinitionsPart styles = main.AddNewPart<StyleDefinitionsPart>();
                Write(styles, $"<w:styles xmlns:w=\"{Wml}\">{stylesXml}</w:styles>");
            }

            if (numberingXml is not null)
            {
                NumberingDefinitionsPart numbering = main.AddNewPart<NumberingDefinitionsPart>();
                Write(numbering, $"<w:numbering xmlns:w=\"{Wml}\">{numberingXml}</w:numbering>");
            }

            extra?.Invoke(main);
        }

        return file.ToArray();
    }

    public static void Write(OpenXmlPart part, string xml)
    {
        using Stream stream = part.GetStream(FileMode.Create);
        using StreamWriter writer = new(stream);
        writer.Write(xml);
    }

    public static async Task<OrdDocument> OpenAsync(byte[] file, ReadOptions? options = null)
    {
        Result<OrdDocument> opened = await OrdDocument.OpenAsync(new MemoryStream(file), options);

        Assert.True(opened.IsSuccess(out OrdDocument? document), opened.Error);

        return document;
    }

    public static async Task<byte[]> SaveAsync(OrdDocument document, SaveOptions? options = null)
    {
        using MemoryStream output = new();
        await document.SaveAsync(output, options);

        return output.ToArray();
    }

    /// <summary>The errors Microsoft's validator finds, as text, one per line; empty when the file is valid.</summary>
    public static string Validate(byte[] file)
    {
        using MemoryStream stream = new(file);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, false);
        OpenXmlValidator validator = new(FileFormatVersions.Microsoft365);
        IEnumerable<ValidationErrorInfo> errors = validator.Validate(document);

        return string.Join('\n', errors.Select(error => $"{error.Part?.Uri} {error.Path?.XPath}: {error.Description}"));
    }

    /// <summary>Fails with every error Microsoft's validator finds.</summary>
    public static void AssertValid(byte[] file)
    {
        string errors = Validate(file);

        Assert.True(errors.Length == 0, errors);
    }

    /// <summary>The body's text as the Open XML SDK reads it, paragraph by paragraph.</summary>
    public static string SdkText(byte[] file)
    {
        using MemoryStream stream = new(file);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, false);
        WordDocument body = document.MainDocumentPart!.Document!;

        return string.Join('\n', body.Body!.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Select(paragraph => paragraph.InnerText));
    }
}
