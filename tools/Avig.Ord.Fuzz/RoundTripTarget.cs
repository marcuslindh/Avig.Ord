using System.Text;

namespace Avig.Ord.Fuzz;

/// <summary>
/// <c>roundtrip</c> — open, save, open gives the same document. One part of the plain or the
/// rich template is mutated, as in <c>docx-part</c>. When the document opens:
///
/// - saved untouched (the parts never read are written byte for byte) and saved after being
///   read (the parts are written from their XML), it reads back with the same plain text and
///   Markdown;
/// - when the input passes Microsoft's schema validator, the saved file does too, also after the
///   changes every file target makes.
/// </summary>
internal static class RoundTripTarget
{
    private static readonly string[] PartNames =
    [
        "word/document.xml",
        "word/styles.xml",
        "word/numbering.xml",
        "word/header1.xml",
        "word/footnotes.xml",
    ];

    public static FuzzTarget Create()
    {
        return new FuzzTarget("roundtrip", "open, save, open gives the same text; a valid input saves as a valid file", RunAsync, Generate);
    }

    private static byte[] Generate(Random random)
    {
        int part = random.Next(PartNames.Length);
        bool rich = random.Next(3) == 0;
        string content = Template.Text(rich, PartNames[part]);

        // Few mutations: the validator invariant only says something about inputs that are still valid.
        if (random.Next(5) > 0)
        {
            content = Inputs.Mutate(random, content, PartTarget.Tokens);
        }

        byte[] text = Encoding.UTF8.GetBytes(content);
        byte[] input = new byte[text.Length + 2];
        input[0] = (byte)part;

        if (rich == true)
        {
            input[0] |= 0x80;
        }

        input[1] = (byte)random.Next(256);
        text.CopyTo(input, 2);

        return input;
    }

    private static async Task RunAsync(byte[] input)
    {
        if (input.Length < 2)
        {
            return;
        }

        bool rich = (input[0] & 0x80) != 0;
        string part = PartNames[(input[0] & 0x7F) % PartNames.Length];
        RevisionView view = (RevisionView)(input[1] % 3);
        bool flat = (input[1] & 4) != 0;
        await using MemoryStream built = Template.Build(Template.Parts(rich), part, input.AsSpan(2));
        await Dump.WriteAsync("input.docx", built);
        byte[] file = built.ToArray();
        ReadOptions options = new() { Revisions = view };

        Result<OrdDocument> untouched = await OrdDocument.OpenAsync(new MemoryStream(file, false), options);

        if (untouched.IsSuccess(out OrdDocument? first) == false)
        {
            Harness.CheckKind(untouched.ErrorKind(), untouched.Error);

            return;
        }

        bool valid = Validator.IsValid(file);
        byte[] savedUntouched = await Harness.SaveAsync(first, new SaveOptions { FlatOpc = flat });

        Result<OrdDocument> opened = await OrdDocument.OpenAsync(new MemoryStream(file, false), options);

        if (opened.IsSuccess(out OrdDocument? document) == false)
        {
            throw new InvariantViolationException($"The same file opened once and not twice: {opened.Error}");
        }

        string text = document.ToPlainText();
        string markdown = document.ToMarkdown();
        byte[] savedRead = await Harness.SaveAsync(document, SaveOptions.Default);

        await CheckSameAsync("saved untouched", savedUntouched, options, text, markdown);
        await CheckSameAsync("saved after reading", savedRead, options, text, markdown);

        if (valid == false)
        {
            return;
        }

        CheckValid("saved after reading", savedRead);

        if (flat == false)
        {
            CheckValid("saved untouched", savedUntouched);
        }

        Harness.Edit(document);
        byte[] edited = await Harness.SaveAsync(document, SaveOptions.Default);
        await Dump.WriteAsync("edited.docx", new MemoryStream(edited));
        CheckValid("saved after the changes", edited);
    }

    private static async Task CheckSameAsync(string how, byte[] saved, ReadOptions options, string text, string markdown)
    {
        Result<OrdDocument> reopened = await OrdDocument.OpenAsync(new MemoryStream(saved, false), options);

        if (reopened.IsSuccess(out OrdDocument? again) == false)
        {
            throw new InvariantViolationException($"A document {how} cannot be opened again: {reopened.Error}");
        }

        string againText = again.ToPlainText();

        if (againText != text)
        {
            throw new InvariantViolationException($"The text changed when {how}:\n  before: {Harness.Shorten(text)}\n  after:  {Harness.Shorten(againText)}");
        }

        string againMarkdown = again.ToMarkdown();

        if (againMarkdown != markdown)
        {
            throw new InvariantViolationException($"The Markdown changed when {how}:\n  before: {Harness.Shorten(markdown)}\n  after:  {Harness.Shorten(againMarkdown)}");
        }
    }

    private static void CheckValid(string how, byte[] saved)
    {
        string errors = Validator.Validate(saved);

        if (errors.Length > 0)
        {
            throw new InvariantViolationException($"The input passes the schema validator, but the document {how} does not:\n{errors}");
        }
    }
}
