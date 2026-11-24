using System.Text.Json;
using Avig.Ord.WordCheck;

namespace Avig.Ord.Tests;

/// <summary>
/// Documents Microsoft Word made itself, through COM (tools/Avig.Ord.WordCheck create): lists
/// of several levels, merged and nested cells, tracked changes, comments, notes, fields,
/// headers and footers, content controls, a text box, pictures and symbols. Beside each is
/// Word's own reading of its text (.json). Avig.Ord must read the same.
/// </summary>
public class WordFilesTests
{
    private static readonly string Folder = Path.Combine(AppContext.BaseDirectory, "WordFiles");

    public static TheoryData<string> Files()
    {
        TheoryData<string> files = [];

        foreach (string file in Directory.GetFiles(Folder, "*.docx").Order(StringComparer.Ordinal))
        {
            files.Add(Path.GetFileName(file));
        }

        return files;
    }

    [Fact]
    public void Every_kind_of_word_file_is_here()
    {
        Assert.Equal(7, Directory.GetFiles(Folder, "*.docx").Length);
        Assert.All(Directory.GetFiles(Folder, "*.docx"), file => Assert.True(File.Exists(Path.ChangeExtension(file, ".json")), file));
    }

    [Theory]
    [MemberData(nameof(Files))]
    public async Task A_file_word_made_reads_as_word_reads_it(string name)
    {
        string path = Path.Combine(Folder, name);
        Report word = JsonSerializer.Deserialize<Report>(await File.ReadAllTextAsync(Path.ChangeExtension(path, ".json")))!;

        Report ord = await OrdSide.ReadAsync(path);

        Assert.Equal(string.Empty, string.Join('\n', OrdSide.Compare(word, ord)));
    }

    [Theory]
    [MemberData(nameof(Files))]
    public async Task A_file_word_made_saves_and_reads_back_the_same(string name)
    {
        string path = Path.Combine(Folder, name);
        OrdDocument document = await OrdSide.OpenAsync(path, RevisionView.Marked);
        string text = document.ToPlainText();

        byte[] saved = await TestFiles.SaveAsync(document);

        Assert.Equal(text, (await TestFiles.OpenAsync(saved, new ReadOptions { Revisions = RevisionView.Marked })).ToPlainText());
        Assert.Equal(string.Empty, TestFiles.Validate(saved));
    }

    [Fact]
    public async Task Word_lists_of_several_levels_have_words_labels()
    {
        Report ord = await OrdSide.ReadAsync(Path.Combine(Folder, "lists.docx"));

        Assert.Equal(["", "1)", "a)", "i)", "b)", "2)"], ord.Body.Take(6).Select(paragraph => paragraph.Label));
    }

    [Fact]
    public async Task Word_tracked_changes_read_in_each_view()
    {
        string path = Path.Combine(Folder, "tracked-changes.docx");

        Assert.Equal("Det var en nygammal text.", (await OrdSide.OpenAsync(path, RevisionView.Marked)).Paragraphs[0].Text);
        Assert.Equal("Det var en ny text.", (await OrdSide.OpenAsync(path, RevisionView.Accepted)).Paragraphs[0].Text);
        Assert.Equal("Det var en gammal text.", (await OrdSide.OpenAsync(path, RevisionView.Original)).Paragraphs[0].Text);
    }
}
