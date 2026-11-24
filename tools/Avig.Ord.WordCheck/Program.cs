using System.Runtime.InteropServices;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace Avig.Ord.WordCheck;

/// <summary>
/// Checks Avig.Ord against Microsoft Word, through COM:
///
/// - <c>create</c> has Word make documents with real Word features and saves them, with Word's
///   own reading of their text, in tests/Avig.Ord.Tests/WordFiles (WordFilesTests reads them);
/// - <c>check</c> writes documents with Avig.Ord in every way it writes them, opens each in Word
///   without repair, and compares Word's text with Avig.Ord's.
/// </summary>
internal static class Program
{
    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    private static async Task<int> Main(string[] args)
    {
        string command = "all";

        if (args.Length > 0)
        {
            command = args[0].ToLowerInvariant();
        }

        string wordFiles = Path.Combine(RepositoryRoot(), "tests", "Avig.Ord.Tests", "WordFiles");
        string output = Path.Combine(Path.GetTempPath(), "Avig.Ord.WordCheck");
        int failures = 0;

        using (WordSession session = WordSession.Start())
        {
            Console.WriteLine($"Word {session.Version}");
            Console.WriteLine();

            if (command is "create" or "all")
            {
                failures += await CreateAsync(session, wordFiles);
            }

            if (command is "check" or "all")
            {
                failures += await CheckAsync(session, output, wordFiles);
            }
        }

        Console.WriteLine();

        if (failures > 0)
        {
            Console.WriteLine($"FAIL: {failures} documents differ or did not open.");

            return 1;
        }

        Console.WriteLine("OK: Word opened every document and read the same text as Avig.Ord.");

        return 0;
    }

    /// <summary>Word makes its documents; each is read back by Word and by Avig.Ord, and Word's reading is kept beside it.</summary>
    private static async Task<int> CreateAsync(WordSession session, string folder)
    {
        Console.WriteLine($"Word makes documents in {folder}");
        int failures = 0;

        foreach ((string path, Report word) in WordCreated.CreateAll(session, folder))
        {
            await File.WriteAllTextAsync(Path.ChangeExtension(path, ".json"), JsonSerializer.Serialize(word, Json).ReplaceLineEndings("\r\n") + "\r\n");
            failures += await CompareAsync(path, word);
        }

        Console.WriteLine();

        return failures;
    }

    /// <summary>Avig.Ord writes its documents, and Word opens each.</summary>
    private static async Task<int> CheckAsync(WordSession session, string folder, string wordFiles)
    {
        Console.WriteLine($"Avig.Ord writes documents in {folder}; Word opens them");
        int failures = 0;

        foreach (string path in await OrdWritten.WriteAllAsync(folder, wordFiles))
        {
            Report word;

            try
            {
                dynamic document = session.Open(path);

                try
                {
                    string name = (string)document.Name;

                    if (string.Equals(name, Path.GetFileName(path), StringComparison.OrdinalIgnoreCase) == false)
                    {
                        Console.WriteLine($"  NOTE {Path.GetFileName(path)}: Word opened it as '{name}'");
                    }

                    word = WordSession.Read(document);
                }
                finally
                {
                    WordSession.Close(document);
                }
            }
            catch (COMException exception)
            {
                failures++;
                Console.WriteLine($"  FAIL {Path.GetFileName(path)}: Word did not open it: {exception.Message}");

                continue;
            }

            failures += await CompareAsync(path, word);
        }

        return failures;
    }

    private static async Task<int> CompareAsync(string path, Report word)
    {
        List<string> differences = OrdSide.Compare(word, await OrdSide.ReadAsync(path));

        if (differences.Count == 0)
        {
            Console.WriteLine($"  OK   {Path.GetFileName(path),-48} {word.Body.Count} paragraphs, {word.Headers.Count + word.Footers.Count} header and footer lines, {word.Footnotes.Count + word.Endnotes.Count} notes, {word.Comments.Count} comments, {word.TextBoxes.Count} text box lines");

            return 0;
        }

        Console.WriteLine($"  DIFF {Path.GetFileName(path)}");

        foreach (string difference in differences)
        {
            Console.WriteLine($"       {difference}");
        }

        return 1;
    }

    /// <summary>The folder with Avig.Ord.slnx, however the tool was started.</summary>
    private static string RepositoryRoot()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);

        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Avig.Ord.slnx")) == true)
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        return Environment.CurrentDirectory;
    }
}
