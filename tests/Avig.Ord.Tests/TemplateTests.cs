using System.Globalization;

namespace Avig.Ord.Tests;

public class TemplateTests
{
    private static readonly CultureInfo Swedish = CultureInfo.GetCultureInfo("sv-SE");

    [Fact]
    public async Task Placeholders_are_filled_with_formats_in_the_culture_and_nested_values()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:t>{{Namn}} ska betala {{Belopp:N2}} kr senast {{Datum:d MMMM yyyy}}. Kontakt: {{Kund.Epost}}</w:t></w:r></w:p>"));
        TemplateValues values = new TemplateValues()
            .Set("Namn", "Anna")
            .Set("Belopp", 1200.5m)
            .Set("Datum", new DateOnly(2026, 12, 24))
            .Set("Kund", new TemplateValues().Set("Epost", "anna@example.com"));

        TemplateResult result = document.FillTemplate(values, new TemplateOptions { Culture = Swedish });

        Assert.Equal("Anna ska betala 1 200,50 kr senast 24 december 2026. Kontakt: anna@example.com", document.ToPlainText());
        Assert.Equal(4, result.Filled);
        Assert.Empty(result.Missing);
    }

    [Fact]
    public async Task A_missing_value_becomes_empty_and_is_listed_with_where_it_is()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:t>Kundnummer: {{Kundnumer}}</w:t></w:r></w:p>"));

        TemplateResult result = document.FillTemplate(new TemplateValues().Set("Kundnummer", 42));

        Assert.Equal("Kundnummer: ", document.ToPlainText());
        MissingValue missing = Assert.Single(result.Missing);
        Assert.Equal("Kundnumer", missing.Name);
        Assert.Contains("word/document.xml", missing.Location, StringComparison.Ordinal);
        Assert.Contains("Kundnummer: {{Kundnumer}}", missing.Location, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_section_over_table_rows_repeats_the_rows_for_each_item()
    {
        const string body = """
            <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
              <w:tr><w:tc><w:p><w:r><w:t>Vara</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Pris</w:t></w:r></w:p></w:tc></w:tr>
              <w:tr><w:tc><w:p><w:r><w:t>{{#Rader}}{{Vara}}</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>{{Pris:0}}{{/Rader}}</w:t></w:r></w:p></w:tc></w:tr>
              <w:tr><w:tc><w:p><w:r><w:t>Summa</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>{{Summa}}</w:t></w:r></w:p></w:tc></w:tr>
            </w:tbl><w:p/>
            """;
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx(body));
        TemplateValues values = new TemplateValues()
            .Set("Rader", new[] { new TemplateValues().Set("Vara", "Te").Set("Pris", 25), new TemplateValues().Set("Vara", "Kaffe").Set("Pris", 30) })
            .Set("Summa", 55);

        TemplateResult result = document.FillTemplate(values, new TemplateOptions { Culture = Swedish });

        Assert.Empty(result.Missing);
        Assert.Equal("Vara\tPris\nTe\t25\nKaffe\t30\nSumma\t55\n", document.ToPlainText());
        Assert.True(document.Tables[0].Rows[2].Cells[1].Paragraphs[0].Runs[0].Format.Bold);
        TestFiles.AssertValid(await TestFiles.SaveAsync(document));
    }

    [Fact]
    public async Task A_section_over_paragraphs_repeats_them_and_drops_the_tag_paragraphs()
    {
        const string body = """
            <w:p><w:r><w:t>Deltagare:</w:t></w:r></w:p>
            <w:p><w:r><w:t>{{#Personer}}</w:t></w:r></w:p>
            <w:p><w:r><w:t>{{Namn}}, {{Roll}}</w:t></w:r></w:p>
            <w:p><w:r><w:t>{{/Personer}}</w:t></w:r></w:p>
            <w:p><w:r><w:t>Slut</w:t></w:r></w:p>
            """;
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx(body));
        TemplateValues values = new TemplateValues().Set("Personer", new[] { new TemplateValues().Set("Namn", "Anna").Set("Roll", "ordförande"), new TemplateValues().Set("Namn", "Bo").Set("Roll", "sekreterare") });

        document.FillTemplate(values);

        Assert.Equal("Deltagare:\nAnna, ordförande\nBo, sekreterare\nSlut", document.ToPlainText());
    }

    [Theory]
    [InlineData(true, "Före\nVisas\nEfter")]
    [InlineData(false, "Före\nAnnars\nEfter")]
    public async Task Conditions_keep_or_drop_what_they_hold(bool condition, string expected)
    {
        const string body = """
            <w:p><w:r><w:t>Före</w:t></w:r></w:p>
            <w:p><w:r><w:t>{{#Villkor}}</w:t></w:r></w:p><w:p><w:r><w:t>Visas</w:t></w:r></w:p><w:p><w:r><w:t>{{/Villkor}}</w:t></w:r></w:p>
            <w:p><w:r><w:t>{{^Villkor}}</w:t></w:r></w:p><w:p><w:r><w:t>Annars</w:t></w:r></w:p><w:p><w:r><w:t>{{/Villkor}}</w:t></w:r></w:p>
            <w:p><w:r><w:t>Efter</w:t></w:r></w:p>
            """;
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx(body));

        document.FillTemplate(new TemplateValues().Set("Villkor", condition));

        Assert.Equal(expected, document.ToPlainText());
    }

    [Fact]
    public async Task A_section_inside_one_paragraph_is_filled_as_text()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:t>Taggar: {{#Taggar}}{{.}}; {{/Taggar}}{{#Admin}}(admin){{/Admin}}{{^Gäst}}!{{/Gäst}}</w:t></w:r></w:p>"));

        string[] tags = ["a", "b"];
        document.FillTemplate(new TemplateValues().Set("Taggar", tags).Set("Admin", true));

        Assert.Equal("Taggar: a; b; (admin)!", document.ToPlainText());
    }

    [Fact]
    public async Task Nested_sections_look_names_up_from_the_inside_out()
    {
        const string body = """
            <w:p><w:r><w:t>{{#Grupper}}</w:t></w:r></w:p>
            <w:p><w:r><w:t>{{Namn}} ({{Firma}}):</w:t></w:r></w:p>
            <w:p><w:r><w:t>{{#Medlemmar}}- {{Namn}}{{/Medlemmar}}</w:t></w:r></w:p>
            <w:p><w:r><w:t>{{/Grupper}}</w:t></w:r></w:p>
            """;
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx(body));
        TemplateValues values = new TemplateValues()
            .Set("Firma", "AB")
            .Set("Grupper", new[]
            {
                new TemplateValues().Set("Namn", "Ekonomi").Set("Medlemmar", new[] { new TemplateValues().Set("Namn", "Anna") }),
                new TemplateValues().Set("Namn", "IT").Set("Medlemmar", Array.Empty<TemplateValues>()),
            });

        document.FillTemplate(values);

        Assert.Equal("Ekonomi (AB):\n- Anna\nIT (AB):\n", document.ToPlainText());
    }

    [Fact]
    public async Task A_replacement_as_value_puts_in_a_table_or_a_picture()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:t>{{Logga}}</w:t></w:r></w:p><w:p><w:r><w:t>{{Priser}}</w:t></w:r></w:p><w:sectPr/>"));
        TemplateValues values = new TemplateValues()
            .Set("Logga", Replacement.Picture(WriterTests.Png, "Logga"))
            .Set("Priser", Replacement.Table(table => table.HeaderRow("Vara", "Pris").Row("Te", "25")));

        document.FillTemplate(values);

        Assert.Single(document.AllPictures());
        Assert.Single(document.Tables);
        TestFiles.AssertValid(await TestFiles.SaveAsync(document));
    }

    [Fact]
    public async Task Merge_fields_are_filled_with_their_switches()
    {
        const string body = """
            <w:p>
              <w:r><w:t xml:space="preserve">Hej </w:t></w:r>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> MERGEFIELD Förnamn \* Upper \* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>«Förnamn»</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
              <w:r><w:t xml:space="preserve">, född </w:t></w:r>
              <w:fldSimple w:instr=" MERGEFIELD Född \@ &quot;yyyy-MM-dd&quot; "><w:r><w:t>«Född»</w:t></w:r></w:fldSimple>
              <w:fldSimple w:instr=" MERGEFIELD Titel \b &quot;titel: &quot; "><w:r><w:t>«Titel»</w:t></w:r></w:fldSimple>
            </w:p>
            """;
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx(body));

        TemplateResult result = document.FillTemplate(new TemplateValues().Set("Förnamn", "Anna").Set("Född", new DateTime(1990, 3, 4)));

        Assert.Equal("Hej ANNA, född 1990-03-04", document.ToPlainText());
        Assert.True(document.Paragraphs[0].Runs.Single(run => run.Text == "ANNA").Format.Bold);
        Assert.Equal("Titel", Assert.Single(result.Missing).Name);
        Assert.Empty(document.Paragraphs[0].Inlines.OfType<Field>());
        TestFiles.AssertValid(await TestFiles.SaveAsync(document));
    }

    [Fact]
    public async Task Content_controls_are_filled_by_tag_and_others_keep_their_content()
    {
        const string body = """
            <w:sdt><w:sdtPr><w:tag w:val="Adress"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Klicka här</w:t></w:r></w:p></w:sdtContent></w:sdt>
            <w:p><w:sdt><w:sdtPr><w:alias w:val="Namn"/></w:sdtPr><w:sdtContent><w:r><w:t>Namn här</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> och </w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="Formulärfält"/></w:sdtPr><w:sdtContent><w:r><w:t>Behålls</w:t></w:r></w:sdtContent></w:sdt></w:p>
            """;
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx(body));

        document.FillTemplate(new TemplateValues().Set("Adress", "Storgatan 1").Set("Namn", "Anna"));

        Assert.Equal("Storgatan 1\nAnna och Behålls", document.ToPlainText());
        TestFiles.AssertValid(await TestFiles.SaveAsync(document));
    }

    [Fact]
    public async Task Headers_footers_and_notes_are_filled_too()
    {
        using MemoryStream file = new();

        await using (DocxWriter writer = await DocxWriter.CreateAsync(file, leaveOpen: true))
        {
            writer.Header(header => header.Paragraph("{{Firma}}"));
            writer.Footer(footer => footer.Paragraph("Sida {{Sida}}"));
            await writer.ParagraphAsync(paragraph => paragraph.Text("Text").Footnote("Källa: {{Källa}}"));
        }

        OrdDocument document = await TestFiles.OpenAsync(file.ToArray());
        TemplateResult result = document.FillTemplate(new TemplateValues().Set("Firma", "AB").Set("Källa", "SCB"));

        Assert.Equal("AB", document.Headers[0].Text);
        Assert.Equal("Källa: SCB", document.Footnotes[0].Text);
        Assert.Equal("Sida", Assert.Single(result.Missing).Name);
    }

    [Fact]
    public async Task A_template_filled_with_tracked_changes_reads_as_the_template_rejected()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:t>Hej {{Namn}}!</w:t></w:r></w:p>"));
        document.TrackChanges("Mall");

        document.FillTemplate(new TemplateValues().Set("Namn", "Anna"));
        byte[] saved = await TestFiles.SaveAsync(document);

        TestFiles.AssertValid(saved);
        Assert.Equal("Hej Anna!", (await TestFiles.OpenAsync(saved)).ToPlainText());
        Assert.Equal("Hej {{Namn}}!", (await TestFiles.OpenAsync(saved, new ReadOptions { Revisions = RevisionView.Original })).ToPlainText());
    }

    [Fact]
    public async Task A_section_without_an_end_is_removed_and_listed()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:t>{{#Öppen}} text {{/Annan}}</w:t></w:r></w:p>"));

        TemplateResult result = document.FillTemplate(new TemplateValues());

        Assert.Equal(" text ", document.ToPlainText());
        Assert.Equal(2, result.Missing.Count);
    }

    [Fact]
    public async Task A_dictionary_works_as_values()
    {
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx("<w:p><w:r><w:t>{{a}}-{{B}}</w:t></w:r></w:p>"));

        document.FillTemplate(new Dictionary<string, object?> { ["A"] = 1, ["b"] = "två" });

        Assert.Equal("1-två", document.ToPlainText());
    }
}
