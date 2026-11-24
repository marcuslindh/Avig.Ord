using System.Runtime.InteropServices;

namespace Avig.Ord.WordCheck;

/// <summary>
/// One hidden Word, driven through COM with late binding. Word is quit when the session is
/// disposed, also after an exception, so that no WINWORD.EXE is left behind.
/// </summary>
internal sealed class WordSession : IDisposable
{
    private const int DoNotSaveChanges = 0;
    private const int AlertsNone = 0;
    private const int ListBullet = 2;
    private const int ListPictureBullet = 6;

    private readonly dynamic word;

    private WordSession(dynamic word)
    {
        this.word = word;
    }

    public dynamic Application => word;

    public string Version => (string)word.Version + " (build " + (string)word.Build + ")";

    /// <summary>Starts Word, hidden, with every alert off: a prompt it cannot show is an error instead.</summary>
    public static WordSession Start()
    {
        Type type = Type.GetTypeFromProgID("Word.Application") ?? throw new InvalidOperationException("Word is not installed: there is no Word.Application.");
        dynamic word = Activator.CreateInstance(type) ?? throw new InvalidOperationException("Word did not start.");
        word.Visible = false;
        word.DisplayAlerts = AlertsNone;

        return new WordSession(word);
    }

    /// <summary>Opens a file read-only, without repairing it and without asking anything.</summary>
    public dynamic Open(string path)
    {
        return word.Documents.Open(
            FileName: path,
            ConfirmConversions: false,
            ReadOnly: true,
            AddToRecentFiles: false,
            Revert: false,
            Visible: true,
            OpenAndRepair: false,
            NoEncodingDialog: true);
    }

    public dynamic NewDocument()
    {
        return word.Documents.Add(Visible: true);
    }

    public static void Close(dynamic document)
    {
        document.Close(SaveChanges: DoNotSaveChanges);
        Marshal.FinalReleaseComObject(document);
    }

    /// <summary>What Word reads in a document: the body with tracked changes as they are, then with them accepted (the document is closed without saving), and the other stories.</summary>
    public static Report Read(dynamic document)
    {
        // Word gives an inline picture as "/" in the text; the package gives nothing. The
        // pictures are counted and taken out of this copy (it is closed without saving).
        int pictures = (int)document.InlineShapes.Count;

        while ((int)document.InlineShapes.Count > 0)
        {
            document.InlineShapes.Item(1).Delete();
        }

        Report report = new()
        {
            Pictures = pictures,
            Body = Body(document),
            Headers = HeadersAndFooters(document, true),
            Footers = HeadersAndFooters(document, false),
            Footnotes = Notes(document.Footnotes),
            Endnotes = Notes(document.Endnotes),
            Comments = Notes(document.Comments),
            TextBoxes = TextBoxes(document),
        };

        document.Revisions.AcceptAll();
        report.Accepted.AddRange(Body(document));

        return report;
    }

    public void Dispose()
    {
        try
        {
            word.Quit(SaveChanges: DoNotSaveChanges);
        }
        catch (COMException)
        {
            // Word already went; nothing is left to quit.
        }

        Marshal.FinalReleaseComObject(word);
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    private static List<ParagraphText> Body(dynamic document)
    {
        List<ParagraphText> paragraphs = [];

        // Word ends the last paragraph of a section but the last with the section break, ;
        // the package has the break in the paragraph's properties, not in its text.
        HashSet<int> sectionEnds = [];
        int sections = (int)document.Sections.Count;

        for (int index = 1; index < sections; index++)
        {
            sectionEnds.Add((int)document.Sections.Item(index).Range.End);
        }

        foreach (dynamic paragraph in document.Paragraphs)
        {
            dynamic range = paragraph.Range;
            string raw = (string)range.Text;

            if (sectionEnds.Contains((int)range.End) == true && raw.EndsWith((char)12))
            {
                raw = raw[..^1];
            }

            string text = WordText.Normalize(raw);

            if (text.Length == 0)
            {
                continue;
            }

            int listType = (int)range.ListFormat.ListType;
            string label = (string)range.ListFormat.ListString;

            if (listType is ListBullet or ListPictureBullet)
            {
                label = "bullet";
            }

            paragraphs.Add(new ParagraphText(label, text));
        }

        return paragraphs;
    }

    private static List<string> HeadersAndFooters(dynamic document, bool headers)
    {
        List<string> lines = [];

        foreach (dynamic section in document.Sections)
        {
            dynamic stories = section.Footers;

            if (headers == true)
            {
                stories = section.Headers;
            }

            for (int index = 1; index <= 3; index++)
            {
                dynamic story = stories.Item(index);

                if ((bool)story.Exists == true)
                {
                    lines.AddRange(WordText.MaskNumbers(WordText.Lines((string)story.Range.Text)));
                }
            }
        }

        return [.. lines.Distinct().Order(StringComparer.Ordinal)];
    }

    private static List<string> Notes(dynamic notes)
    {
        List<string> texts = [];

        foreach (dynamic note in notes)
        {
            texts.Add(string.Join('\n', WordText.Lines((string)note.Range.Text)));
        }

        return texts;
    }

    private static List<string> TextBoxes(dynamic document)
    {
        List<string> lines = [];

        foreach (dynamic shape in document.Shapes)
        {
            try
            {
                if ((int)shape.TextFrame.HasText != 0)
                {
                    lines.AddRange(WordText.Lines((string)shape.TextFrame.TextRange.Text));
                }
            }
            catch (COMException)
            {
                // A picture or a line has no text frame.
            }
        }

        return lines;
    }
}
