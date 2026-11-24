using System.Runtime.InteropServices;

namespace Avig.Ord.WordCheck;

/// <summary>
/// Documents Word itself makes, through COM, with what real documents have: lists of several
/// levels, merged and nested table cells, tracked changes, comments, notes, fields, headers and
/// footers, content controls, a text box and pictures. They are saved as .docx for the tests.
/// </summary>
internal static class WordCreated
{
    private const int FormatXmlDocument = 12;
    private const int CollapseEnd = 0;
    private const int BulletGallery = 1;
    private const int OutlineNumberGallery = 3;
    private const int ListApplyToWholeList = 0;
    private const int Word10ListBehavior = 2;
    private const int FieldEmpty = -1;
    private const int FieldPage = 33;
    private const int FieldNumPages = 26;
    private const int PageBreak = 7;
    private const int ContentControlRichText = 0;
    private const int ContentControlText = 1;
    private const int ContentControlDropdown = 4;
    private const int ContentControlCheckBox = 8;
    private const int TextOrientationHorizontal = 1;
    private const int HeaderPrimary = 1;

    /// <summary>The documents, by file name, and how to make each.</summary>
    public static readonly (string Name, Action<WordSession, dynamic, string> Build)[] Documents =
    [
        ("lists.docx", (session, document, folder) => Lists(session, document)),
        ("tables.docx", (session, document, folder) => Tables(document)),
        ("tracked-changes.docx", (session, document, folder) => TrackedChanges(document)),
        ("comments-and-notes.docx", (session, document, folder) => CommentsAndNotes(document)),
        ("fields-headers-footers.docx", (session, document, folder) => FieldsHeadersFooters(document)),
        ("content-controls-text-box.docx", (session, document, folder) => ContentControlsAndTextBox(document)),
        ("pictures-symbols.docx", (session, document, folder) => PicturesAndSymbols(document, folder)),
    ];

    /// <summary>Makes every document in <paramref name="folder"/> and gives Word's own reading of each.</summary>
    public static List<(string Path, Report Word)> CreateAll(WordSession session, string folder)
    {
        Directory.CreateDirectory(folder);
        List<(string Path, Report Word)> created = [];

        foreach ((string name, Action<WordSession, dynamic, string> build) in Documents)
        {
            string path = Path.Combine(folder, name);
            dynamic document = session.NewDocument();

            try
            {
                build(session, document, folder);
                document.SaveAs2(FileName: path, FileFormat: FormatXmlDocument, AddToRecentFiles: false);
            }
            finally
            {
                WordSession.Close(document);
            }

            // Read back from the saved file, as the tests will.
            dynamic reopened = session.Open(path);

            try
            {
                created.Add((path, WordSession.Read(reopened)));
            }
            finally
            {
                WordSession.Close(reopened);
            }
        }

        return created;
    }

    /// <summary>A new paragraph at the end with this text; the range of the text.</summary>
    private static dynamic Append(dynamic document, string text)
    {
        dynamic end = document.Content;
        end.Collapse(CollapseEnd);
        end.InsertParagraphAfter();
        dynamic range = document.Content;
        range.Collapse(CollapseEnd);
        range.Text = text;

        return range;
    }

    /// <summary>The range of the last paragraph's text, collapsed at its end, for a field or a control.</summary>
    private static dynamic End(dynamic document)
    {
        dynamic range = document.Content;
        range.Collapse(CollapseEnd);

        return range;
    }

    private static void Lists(WordSession session, dynamic document)
    {
        document.Content.Text = "Listor av Word";
        dynamic first = Append(document, "Kapitel ett");
        Append(document, "Avsnitt ett och ett");
        Append(document, "Stycke ett och ett och ett");
        Append(document, "Avsnitt ett och två");
        dynamic last = Append(document, "Kapitel två");
        dynamic numbered = document.Range(first.Start, last.End);
        dynamic outline = session.Application.ListGalleries.Item(OutlineNumberGallery).ListTemplates.Item(1);
        numbered.ListFormat.ApplyListTemplateWithLevel(ListTemplate: outline, ContinuePreviousList: false, ApplyTo: ListApplyToWholeList, DefaultListBehavior: Word10ListBehavior);
        SetLevel(document, 3, 2);
        SetLevel(document, 4, 3);
        SetLevel(document, 5, 2);

        Append(document, "Vanligt stycke mellan listorna");
        dynamic bulletStart = Append(document, "Punkt ett");
        Append(document, "Underpunkt");
        dynamic bulletEnd = Append(document, "Punkt två");
        dynamic bullets = document.Range(bulletStart.Start, bulletEnd.End);
        dynamic bullet = session.Application.ListGalleries.Item(BulletGallery).ListTemplates.Item(1);
        bullets.ListFormat.ApplyListTemplateWithLevel(ListTemplate: bullet, ContinuePreviousList: false, ApplyTo: ListApplyToWholeList, DefaultListBehavior: Word10ListBehavior);
        SetLevel(document, 9, 2);
    }

    private static void SetLevel(dynamic document, int paragraph, int level)
    {
        document.Paragraphs.Item(paragraph).Range.ListFormat.ListLevelNumber = level;
    }

    private static void Tables(dynamic document)
    {
        document.Content.Text = "Tabeller av Word";
        Append(document, "Före tabellen");
        dynamic range = End(document);
        range.InsertParagraphAfter();
        range = End(document);
        dynamic table = document.Tables.Add(range, 4, 3);
        table.Borders.Enable = true;
        string[,] cells =
        {
            { "Namn", "Belopp", "Not" },
            { "Hyra", "8 000", "Hela" },
            { "El", "1 200", string.Empty },
            { "Summa", string.Empty, "9 200" },
        };

        for (int row = 0; row < 4; row++)
        {
            for (int column = 0; column < 3; column++)
            {
                table.Cell(row + 1, column + 1).Range.Text = cells[row, column];
            }
        }

        // Vertical first, while the cell numbers are still those of the grid.
        table.Cell(2, 3).Merge(table.Cell(3, 3));
        table.Cell(4, 1).Merge(table.Cell(4, 2));
        dynamic nested = table.Cell(3, 1).Tables.Add(table.Cell(3, 1).Range, 2, 2);
        nested.Cell(1, 1).Range.Text = "a";
        nested.Cell(1, 2).Range.Text = "b";
        nested.Cell(2, 1).Range.Text = "c";
        nested.Cell(2, 2).Range.Text = "d";
        Append(document, "Efter tabellen");
    }

    private static void TrackedChanges(dynamic document)
    {
        document.Content.Text = "Det var en gammal text.";
        Append(document, "Ett stycke som tas bort.");
        Append(document, "Ett stycke som står kvar.");
        document.TrackRevisions = true;
        dynamic found = document.Content;
        found.Find.Execute(FindText: "gammal");
        found.Text = "ny";
        document.Paragraphs.Item(2).Range.Delete();
        dynamic end = Append(document, "Ett nytt stycke.");
        end.InsertAfter(" Med mer text.");
        document.TrackRevisions = false;
    }

    private static void CommentsAndNotes(dynamic document)
    {
        document.Content.Text = "Kommenterad text.";
        document.Comments.Add(document.Paragraphs.Item(1).Range, "Stämmer detta?");
        dynamic second = Append(document, "Text med fotnot");
        dynamic mark = End(document);
        document.Footnotes.Add(Range: mark, Text: "Källa: SCB.");
        Append(document, "Text med slutnot");
        mark = End(document);
        document.Endnotes.Add(Range: mark, Text: "Slutnoten.");
        _ = second;
    }

    private static void FieldsHeadersFooters(dynamic document)
    {
        document.Content.Text = "Fält av Word";
        Append(document, "Kära ");
        document.Fields.Add(Range: End(document), Type: FieldEmpty, Text: "MERGEFIELD Förnamn", PreserveFormatting: false);
        Append(document, "Länk: ");
        document.Hyperlinks.Add(Anchor: End(document), Address: "https://example.com/", TextToDisplay: "Läs mer");
        Append(document, "Sidan är ");
        document.Fields.Add(Range: End(document), Type: FieldPage);
        Append(document, "Före sidbrytningen");
        End(document).InsertBreak(PageBreak);
        Append(document, "Efter sidbrytningen");

        dynamic header = document.Sections.Item(1).Headers.Item(HeaderPrimary).Range;
        header.Text = "Sidhuvud av Word";
        dynamic footer = document.Sections.Item(1).Footers.Item(HeaderPrimary).Range;
        footer.Text = "Sida ";
        document.Fields.Add(Range: StoryEnd(document), Type: FieldPage);
        StoryEnd(document).InsertAfter(" av ");
        document.Fields.Add(Range: StoryEnd(document), Type: FieldNumPages);
    }

    /// <summary>The end of the first footer's text, before its last paragraph mark.</summary>
    private static dynamic StoryEnd(dynamic document)
    {
        dynamic range = document.Sections.Item(1).Footers.Item(HeaderPrimary).Range;
        range.MoveEnd(1, -1);
        range.Collapse(CollapseEnd);

        return range;
    }

    private static void ContentControlsAndTextBox(dynamic document)
    {
        document.Content.Text = "Innehållskontroller av Word";
        Append(document, "Namn: ");
        dynamic text = End(document).ContentControls.Add(ContentControlText);
        text.Tag = "namn";
        text.Title = "Namn";
        text.Range.Text = "Anna";
        Append(document, "Val: ");
        dynamic dropdown = End(document).ContentControls.Add(ContentControlDropdown);
        dropdown.Tag = "val";
        dropdown.DropdownListEntries.Add("Ja", "ja");
        dropdown.DropdownListEntries.Add("Nej", "nej");
        dropdown.DropdownListEntries.Item(2).Select();
        Append(document, "Kryss: ");

        try
        {
            dynamic check = End(document).ContentControls.Add(ContentControlCheckBox);
            check.Tag = "kryss";
            check.Checked = true;
        }
        catch (COMException)
        {
            // A Word before 2010 has no check box control.
        }

        dynamic block = Append(document, "Ett stycke i en blockkontroll");
        dynamic rich = document.Range(block.Start, block.End).ContentControls.Add(ContentControlRichText);
        rich.Tag = "block";
        dynamic anchor = Append(document, "Stycket med textrutan");
        dynamic box = document.Shapes.AddTextbox(Orientation: TextOrientationHorizontal, Left: 100, Top: 100, Width: 200, Height: 50, Anchor: anchor);
        box.TextFrame.TextRange.Text = "I rutan";
        Append(document, "Slut");
    }

    private static void PicturesAndSymbols(dynamic document, string folder)
    {
        document.Content.Text = "Bilder av Word";
        string png = Path.Combine(Path.GetTempPath(), "avig-ord-wordcheck.png");
        File.WriteAllBytes(png, OrdWritten.Png);
        Append(document, "En bild: ");
        dynamic picture = document.InlineShapes.AddPicture(FileName: png, LinkToFile: false, SaveWithDocument: true, Range: End(document));
        picture.AlternativeText = "En röd punkt";

        string svg = Path.Combine(Path.GetTempPath(), "avig-ord-wordcheck.svg");
        File.WriteAllText(svg, OrdWritten.Svg);
        Append(document, "En SVG: ");

        try
        {
            document.InlineShapes.AddPicture(FileName: svg, LinkToFile: false, SaveWithDocument: true, Range: End(document));
        }
        catch (COMException)
        {
            // A Word without SVG support.
            End(document).InsertAfter("(ingen SVG)");
        }

        Append(document, "Symbol: ");
        End(document).InsertSymbol(CharacterNumber: unchecked((short)0xF04A), Font: "Wingdings", Unicode: true);
        Append(document, "Hårt bindestreck: ett" + (char)30 + "två");
        Append(document, "Tabb:\tefter, radbrytning:" + (char)11 + "efter");
        _ = folder;
    }
}
