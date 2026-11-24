using System.Text;
using System.Text.RegularExpressions;

namespace Avig.Ord.Tests;

public class ReplaceTests
{
    /// <summary>"Hej {{Namn}}!" as Word often stores it: split over runs by spell-check marks and formatting.</summary>
    private const string SplitBody = """
        <w:p>
          <w:r><w:t xml:space="preserve">Hej {{</w:t></w:r>
          <w:proofErr w:type="spellStart"/>
          <w:r><w:rPr><w:b/></w:rPr><w:t>Na</w:t></w:r>
          <w:r><w:rPr><w:i/></w:rPr><w:t>mn}}</w:t></w:r>
          <w:proofErr w:type="spellEnd"/>
          <w:r><w:t>!</w:t></w:r>
        </w:p>
        """;

    [Fact]
    public async Task Text_split_over_runs_is_found_and_replaced_in_the_format_it_starts_in()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx(SplitBody));

        Assert.Single(document.Find("{{Namn}}"));
        Assert.Equal(1, document.Replace("{{Namn}}", "Anna"));

        Paragraph paragraph = document.Paragraphs[0];
        Assert.Equal("Hej Anna!", paragraph.Text);
        Assert.False(paragraph.Runs.Single(run => run.Text == "Anna").Format.Bold == true);
        TestFiles.AssertValid(await TestFiles.SaveAsync(document));
    }

    [Fact]
    public async Task A_match_inside_a_run_keeps_the_text_around_it()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>abcabc</w:t></w:r></w:p>"));

        Assert.Equal(2, document.Replace("b", "XY"));

        Assert.Equal("aXYcaXYc", document.ToPlainText());
        Assert.All(document.Paragraphs[0].Runs, run => Assert.True(run.Format.Bold));
    }

    [Fact]
    public async Task Case_whole_words_and_regex_groups_are_honoured()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:t>Katt katten KATT kattunge 2025-11-24</w:t></w:r></w:p>"));

        Assert.Equal(2, document.Replace("katt", "hund", new SearchOptions { MatchCase = false, WholeWord = true }));
        Assert.Equal(1, document.Replace(new Regex(@"(\d{4})-(\d{2})-(\d{2})"), "$3/$2 $1"));

        Assert.Equal("hund katten hund kattunge 24/11 2025", document.ToPlainText());
    }

    [Fact]
    public async Task Headers_footers_and_notes_are_searched_unless_the_scope_says_not()
    {
        byte[] file = await WriterTests.EveryFeatureAsync(false);
        OrdDocument document = await TestFiles.OpenAsync(file);

        Assert.Equal(0, document.Replace("Sidhuvud", "Huvud", new SearchOptions { Scope = SearchScope.Body }));
        Assert.Equal(1, document.Replace("Sidhuvud", "Huvud", new SearchOptions { Scope = SearchScope.HeadersAndFooters }));
        Assert.Equal(1, document.Replace("SCB", "Statistiska centralbyrån"));

        Assert.Equal("Huvud", document.Headers[0].Text);
        Assert.Equal("Källa: Statistiska centralbyrån.", document.Footnotes[0].Text);
    }

    [Fact]
    public async Task A_match_can_be_replaced_with_a_table_or_paragraphs()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:t>Före</w:t></w:r></w:p><w:p><w:r><w:t>{{Tabell}}</w:t></w:r></w:p><w:p><w:r><w:t>Text {{Lista}} efter</w:t></w:r></w:p><w:sectPr/>"));

        document.Replace("{{Tabell}}", Replacement.Table(table => table.HeaderRow("A", "B").Row("1", "2")));
        document.Replace("{{Lista}}", Replacement.Paragraphs(paragraph => paragraph.Bullet().Text("ett"), paragraph => paragraph.Bullet().Text("två")));

        Assert.Equal("Före\nA\tB\n1\t2\nText \n• ett\n• två\n efter", document.ToPlainText());
        Assert.Single(document.Tables);
        TestFiles.AssertValid(await TestFiles.SaveAsync(document));
    }

    [Fact]
    public async Task A_match_can_be_replaced_with_a_picture()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:t>Logga: [LOGO]</w:t></w:r></w:p>"));

        document.Replace("[LOGO]", Replacement.Picture(WriterTests.Png, "Logga"));

        Assert.Equal("Logga: ", document.ToPlainText());
        Picture picture = Assert.Single(document.AllPictures());
        Assert.Equal("Logga", picture.Description);
        TestFiles.AssertValid(await TestFiles.SaveAsync(document));
    }

    [Fact]
    public async Task A_replacement_in_a_style_of_its_own_keeps_the_rest_of_the_format()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:rPr><w:i/><w:sz w:val=\"28\"/></w:rPr><w:t>Datum: X</w:t></w:r></w:p>"));

        document.Replace("X", Replacement.Text("24 november", style => style.Bold().Color("C00000")));

        Run run = document.Paragraphs[0].Runs.Single(candidate => candidate.Text == "24 november");
        Assert.True(run.Format.Bold);
        Assert.True(run.Format.Italic);
        Assert.Equal(Length.Pt(14), run.Format.Size);
        TestFiles.AssertValid(await TestFiles.SaveAsync(document));
    }

    [Fact]
    public async Task Tracked_replacements_read_as_the_old_text_rejected_and_the_new_accepted()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx(SplitBody));
        document.TrackChanges("Ord", new DateTimeOffset(2026, 11, 24, 12, 0, 0, TimeSpan.Zero));

        document.Replace("{{Namn}}", "Anna");
        byte[] saved = await TestFiles.SaveAsync(document);

        TestFiles.AssertValid(saved);
        Assert.Equal("Hej Anna!", (await TestFiles.OpenAsync(saved)).ToPlainText());
        Assert.Equal("Hej {{Namn}}!", (await TestFiles.OpenAsync(saved, new ReadOptions { Revisions = RevisionView.Original })).ToPlainText());
        string xml = Encoding.UTF8.GetString(await ReaderTests.PartAsync(saved, "word/document.xml"));
        Assert.Contains("w:author=\"Ord\"", xml, StringComparison.Ordinal);
        Assert.Contains("2026-11-24T12:00:00Z", xml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Tracked_edits_mark_what_they_add_and_remove()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:t>Första</w:t></w:r></w:p><w:p><w:r><w:t>Andra</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Rad</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>"));
        document.TrackChanges("Ord");

        document.Paragraphs[0].SetText("Ny första");
        document.Paragraphs[1].Remove();
        document.Tables[0].AddRow("Ny rad");
        document.AddParagraph("Sist");
        byte[] saved = await TestFiles.SaveAsync(document);

        TestFiles.AssertValid(saved);
        Assert.Equal("Ny första\nRad\nNy rad\n\nSist", (await TestFiles.OpenAsync(saved)).ToPlainText());
        Assert.Equal("Första\nAndra\nRad\n", (await TestFiles.OpenAsync(saved, new ReadOptions { Revisions = RevisionView.Original })).ToPlainText());

        OrdDocument accepted = await TestFiles.OpenAsync(saved);
        accepted.AcceptAllRevisions();
        Assert.Equal("Ny första\nRad\nNy rad\n\nSist", accepted.ToPlainText());
    }

    [Fact]
    public async Task A_field_instruction_is_not_searched()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>"));

        Assert.Empty(document.Find("PAGE"));
        Assert.Single(document.Find("3"));
    }

    [Fact]
    public async Task Editing_a_new_document_builds_what_was_asked_for()
    {
        OrdDocument document = OrdDocument.Create();
        Paragraph heading = document.AddHeading("Protokoll", 1);
        Paragraph first = document.AddParagraph("Närvarande");
        first.InsertParagraphAfter("Ärenden");
        first.AppendText(": alla", style => style.Italic());
        Table table = document.AddTable(builder => builder.HeaderRow("Punkt", "Beslut").Row("1", "Ja"));
        TableRow copy = table.Rows[1].InsertCopyAfter();
        copy.Cells[0].SetText("2");
        table.Rows[1].Remove();
        heading.SetStyle("Title");

        byte[] saved = await TestFiles.SaveAsync(document);

        TestFiles.AssertValid(saved);
        OrdDocument reopened = await TestFiles.OpenAsync(saved);
        Assert.Equal("Protokoll\nNärvarande: alla\nÄrenden\nPunkt\tBeslut\n2\tJa\n", reopened.ToPlainText());
        Assert.Equal("Title", reopened.Paragraphs[0].StyleId);
    }
}
