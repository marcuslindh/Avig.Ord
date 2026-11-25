using System.Runtime.InteropServices;

namespace Avig.Ord.WordCheck;

/// <summary>
/// Word makes a document with active content of the kinds it stores itself: an embedded OLE package, a linked picture, an INCLUDEPICTURE field and a template attached
/// from a file. Avig.Ord has to find each, take them out, and Word has to open the cleaned file
/// without repairing it and without any of it left. DDE is left out, as Word runs a DDE field when it
/// is added, and so are ActiveX controls, as adding one waits for a security prompt; the tests
/// cover both with files from the Open XML SDK. Everything is made in a temporary folder and nothing of it is kept in the
/// repository, since the links point to paths on this machine.
/// </summary>
internal static class ActiveCheck
{
    private const int FormatXmlDocument = 12;
    private const int FormatXmlTemplate = 14;
    private const int CollapseEnd = 0;
    private const int FieldEmpty = -1;
    private const int InlineShapeEmbeddedOle = 1;
    private const int InlineShapeLinkedOle = 2;
    private const int InlineShapeLinkedPicture = 4;
    private const int InlineShapeOleControl = 5;

    public static async Task<int> RunAsync(WordSession session, string folder)
    {
        Directory.CreateDirectory(folder);
        string original = Path.Combine(folder, "active-content.docx");
        string cleaned = Path.Combine(folder, "active-content-cleaned.docx");
        Console.WriteLine($"Word makes a document with active content in {folder}");

        Make(session, folder, original);

        Result<OrdDocument> opened = await OrdDocument.OpenAsync(original);

        if (opened.IsFailure(out string? error, out _, out OrdDocument? document) == true)
        {
            Console.WriteLine($"  FAIL Avig.Ord did not open Word's document: {error}");

            return 1;
        }

        ActiveContent expected = ActiveContent.EmbeddedObjects | ActiveContent.ActiveFields | ActiveContent.ExternalReferences;
        int failures = 0;
        Console.WriteLine($"  Avig.Ord finds: {document.ActiveContent}");

        foreach (Avig.OpenXml.OpcExternalRelationship reference in document.ExternalReferences)
        {
            Console.WriteLine($"    {reference.SourcePart}: {reference.Kind} -> {reference.Target}");
        }

        if (document.ActiveContent.HasFlag(expected) == false)
        {
            Console.WriteLine($"  FAIL expected at least {expected}");
            failures++;
        }

        document.RemoveActiveContent();
        Result saved = await document.SaveAsync(cleaned);

        if (saved.IsFailure(out error, out _) == true)
        {
            Console.WriteLine($"  FAIL the cleaned document was not saved: {error}");

            return failures + 1;
        }

        Result<OrdDocument> again = await OrdDocument.OpenAsync(cleaned);

        if (again.IsSuccess(out OrdDocument? reopened) == false || reopened.ActiveContent != ActiveContent.None)
        {
            Console.WriteLine("  FAIL the cleaned document still has active content for Avig.Ord");
            failures++;
        }

        failures += InWord(session, cleaned);

        return failures;
    }

    private static void Make(WordSession session, string folder, string path)
    {
        string picture = Path.Combine(folder, "linked.png");
        string payload = Path.Combine(folder, "payload.txt");
        string template = Path.Combine(folder, "attached.dotx");
        File.WriteAllBytes(picture, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEUlEQVR4nGP4z8DwnwEKAA34A/1q6yb3AAAAAElFTkSuQmCC"));
        File.WriteAllText(payload, "A file embedded as an OLE package, as a payload would be.");

        dynamic blank = session.NewDocument();

        try
        {
            blank.SaveAs2(FileName: template, FileFormat: FormatXmlTemplate, AddToRecentFiles: false);
        }
        finally
        {
            WordSession.Close(blank);
        }

        dynamic document = session.NewDocument();

        try
        {
            document.Content.Text = "Före";
            document.InlineShapes.AddOLEObject(FileName: payload, LinkToFile: false, DisplayAsIcon: true, Range: End(document));
            document.InlineShapes.AddPicture(FileName: picture, LinkToFile: true, SaveWithDocument: false, Range: End(document));
            document.Fields.Add(End(document), FieldEmpty, $"INCLUDEPICTURE \"{picture.Replace("\\", "\\\\", StringComparison.Ordinal)}\" \\d", false);
            dynamic end = End(document);
            end.InsertAfter("Efter");
            document.AttachedTemplate = template;
            document.SaveAs2(FileName: path, FileFormat: FormatXmlDocument, AddToRecentFiles: false);
        }
        finally
        {
            WordSession.Close(document);
        }
    }

    /// <summary>Word opens the cleaned file without repair; no control, object, linked picture or link field is left, and the text around them is.</summary>
    private static int InWord(WordSession session, string path)
    {
        dynamic document;

        try
        {
            document = session.Open(path);
        }
        catch (COMException exception)
        {
            Console.WriteLine($"  FAIL Word did not open the cleaned document: {exception.Message}");

            return 1;
        }

        try
        {
            int failures = 0;
            string name = (string)document.Name;

            if (string.Equals(name, Path.GetFileName(path), StringComparison.OrdinalIgnoreCase) == false)
            {
                Console.WriteLine($"  FAIL Word opened it as '{name}', which means it repaired it");
                failures++;
            }

            int active = 0;

            foreach (dynamic shape in document.InlineShapes)
            {
                int type = (int)shape.Type;

                if (type is InlineShapeEmbeddedOle or InlineShapeLinkedOle or InlineShapeLinkedPicture or InlineShapeOleControl)
                {
                    active++;
                }
            }

            int fields = (int)document.Fields.Count;
            string template = (string)document.AttachedTemplate.Name;
            string text = (string)document.Content.Text;
            Console.WriteLine($"  Word opens the cleaned file: {active} objects or controls, {fields} fields, template '{template}'");

            if (active > 0 || fields > 0)
            {
                Console.WriteLine("  FAIL Word still sees active content in the cleaned file");
                failures++;
            }

            if (template.EndsWith("attached.dotx", StringComparison.OrdinalIgnoreCase) == true)
            {
                Console.WriteLine("  FAIL the attached template is still there");
                failures++;
            }

            if (text.Contains("Före", StringComparison.Ordinal) == false || text.Contains("Efter", StringComparison.Ordinal) == false)
            {
                Console.WriteLine("  FAIL the text around the active content is gone");
                failures++;
            }

            if (failures == 0)
            {
                Console.WriteLine("  OK");
            }

            return failures;
        }
        finally
        {
            WordSession.Close(document);
        }
    }

    private static dynamic End(dynamic document)
    {
        dynamic range = document.Content;
        range.Collapse(CollapseEnd);

        return range;
    }
}
