using System.Text;
using System.Xml.Linq;

namespace Avig.Ord.Fuzz;

/// <summary>
/// Targets whose input is the content of a w:body, put into the rich template in place of its
/// own. The generators make it from <see cref="BodyGenerator"/>'s grammar, so that it has the
/// structure Word writes; libFuzzer mutates it from there.
/// </summary>
internal static class BodyTargets
{
    private static readonly RevisionView[] Views = [RevisionView.Accepted, RevisionView.Original, RevisionView.Marked];

    /// <summary><c>render</c> — a body from the grammar, now and then mutated: both readers, every view, changed and saved, and nothing throws but the documented refusals.</summary>
    public static FuzzTarget Render()
    {
        return new FuzzTarget("render", "a generated body: ToPlainText, ToMarkdown and everything else never throw, in every view", RunRenderAsync, GenerateMutated);
    }

    /// <summary><c>events</c> — <see cref="DocxReader.ReadEventsAsync"/> on a generated, mutated body never throws but the documented refusals.</summary>
    public static FuzzTarget Events()
    {
        return new FuzzTarget("events", "a generated, mutated body: ReadEventsAsync never throws but the documented refusals", RunEventsAsync, GenerateMutated);
    }

    /// <summary><c>reader-model</c> — the streaming reader and the document model read a body the same.</summary>
    public static FuzzTarget ReaderModel()
    {
        return new FuzzTarget("reader-model", "a generated body: DocxReader blocks, events and text equal OrdDocument's", RunReaderModelAsync, Generate);
    }

    /// <summary><c>revisions</c> — accepting every change gives the Accepted view's text, rejecting them the Original view's.</summary>
    public static FuzzTarget Revisions()
    {
        return new FuzzTarget("revisions", "a generated body: AcceptAll gives the Accepted view, RejectAll the Original view", RunRevisionsAsync, Generate);
    }

    private static byte[] Generate(Random random)
    {
        return Encoding.UTF8.GetBytes(BodyGenerator.Body(random));
    }

    private static byte[] GenerateMutated(Random random)
    {
        string body = BodyGenerator.Body(random);

        if (random.Next(3) > 0)
        {
            body = Inputs.Mutate(random, body, PartTarget.Tokens);
        }

        return Encoding.UTF8.GetBytes(body);
    }

    private static byte[] File(byte[] input)
    {
        using MemoryStream file = Template.WithBody(true, Encoding.UTF8.GetString(input));

        return file.ToArray();
    }

    private static async Task RunRenderAsync(byte[] input)
    {
        byte[] file = File(input);
        await Dump.WriteAsync("input.docx", new MemoryStream(file));

        for (int view = 0; view < Views.Length; view++)
        {
            await Harness.RunFileAsync(file, view);
        }
    }

    private static async Task RunEventsAsync(byte[] input)
    {
        byte[] file = File(input);

        foreach (RevisionView view in Views)
        {
            await Harness.StreamAsync(file, new ReadOptions { Revisions = view }, async reader =>
            {
                int depth = 0;

                await foreach (DocumentEvent item in reader.ReadEventsAsync())
                {
                    if (item.Kind == DocumentEventKind.TableStart)
                    {
                        depth++;
                    }
                    else if (item.Kind == DocumentEventKind.TableEnd)
                    {
                        depth--;
                    }

                    if (item.TableDepth < 0 || (item.Kind == DocumentEventKind.Text && item.Text.Length == 0))
                    {
                        throw new InvariantViolationException($"A malformed event: {item}");
                    }
                }
            });
        }
    }

    private static async Task RunReaderModelAsync(byte[] input)
    {
        byte[] file = File(input);
        await Dump.WriteAsync("input.docx", new MemoryStream(file));

        foreach (RevisionView view in Views)
        {
            ReadOptions options = new() { Revisions = view };

            Result<OrdDocument> opened = await OrdDocument.OpenAsync(new MemoryStream(file, false), options);

            if (opened.IsSuccess(out OrdDocument? document) == false)
            {
                throw new InvariantViolationException($"A generated body is refused: {opened.Error}");
            }

            IReadOnlyList<Block> blocks = document.Blocks;
            List<string> modelBlocks = [.. blocks.Select(block => block.Text)];
            List<string> modelLabels = [.. OrdDocument.AllParagraphs(blocks).Select(paragraph => paragraph.List?.Label ?? "-")];
            List<string> modelParagraphs = [.. OrdDocument.AllParagraphs(blocks).Select(paragraph => paragraph.Text)];
            string modelText = document.ToPlainText();
            string modelMarkdown = document.ToMarkdown();

            List<Block> streamed = [];
            await Harness.StreamAsync(file, options, async reader =>
            {
                await foreach (Block block in reader.ReadBlocksAsync())
                {
                    streamed.Add(block);
                }
            });

            Compare(view, "block texts", modelBlocks, [.. streamed.Select(block => block.Text)]);
            Compare(view, "list labels", modelLabels, [.. OrdDocument.AllParagraphs(streamed).Select(paragraph => paragraph.List?.Label ?? "-")]);

            string streamedText = string.Empty;
            await Harness.StreamAsync(file, options, async reader => streamedText = await reader.ReadPlainTextAsync());
            Compare(view, "plain text", [modelText], [streamedText]);

            string streamedMarkdown = string.Empty;
            await Harness.StreamAsync(file, options, async reader =>
            {
                await using StringWriter markdown = new();
                await reader.WriteMarkdownAsync(markdown);
                streamedMarkdown = markdown.ToString();
            });

            Compare(view, "Markdown", [modelMarkdown], [streamedMarkdown]);

            List<string> eventParagraphs = [];
            await Harness.StreamAsync(file, options, async reader => eventParagraphs = await EventParagraphsAsync(reader));
            Compare(view, "paragraph texts from events", modelParagraphs, eventParagraphs);
        }
    }

    /// <summary>The paragraphs' text as the events give it, as the test The_event_reader_gives_the_same_text_as_the_paragraphs builds it.</summary>
    private static async Task<List<string>> EventParagraphsAsync(DocxReader reader)
    {
        List<string> paragraphs = [];
        StringBuilder current = new();

        await foreach (DocumentEvent item in reader.ReadEventsAsync())
        {
            switch (item.Kind)
            {
                case DocumentEventKind.ParagraphStart:
                    current.Clear();

                    break;

                case DocumentEventKind.Text:
                    current.Append(item.Text);

                    break;

                case DocumentEventKind.Break:
                    current.Append('\n');

                    break;

                case DocumentEventKind.ParagraphEnd:
                    paragraphs.Add(current.ToString());

                    break;

                default:
                    break;
            }
        }

        return paragraphs;
    }

    private static async Task RunRevisionsAsync(byte[] input)
    {
        byte[] file = File(input);
        await Dump.WriteAsync("input.docx", new MemoryStream(file));
        string accepted = BodyText(await OpenAsync(file, RevisionView.Accepted));
        string original = BodyText(await OpenAsync(file, RevisionView.Original));
        bool comparable = JoinsKeepFields(await OpenAsync(file, RevisionView.Marked));

        foreach (RevisionView view in Views)
        {
            OrdDocument document = await OpenAsync(file, view);
            document.AcceptAllRevisions();
            OrdDocument again = await OpenAsync(await Harness.SaveAsync(document, SaveOptions.Default), RevisionView.Marked);

            if (comparable == true)
            {
                Compare(view, "text after AcceptAllRevisions and the Accepted view's", [accepted], [BodyText(document)]);
                Compare(view, "text after AcceptAllRevisions and saving, in the Marked view", [accepted], [BodyText(again)]);
            }

            document = await OpenAsync(file, view);
            document.RejectAllRevisions();

            if (comparable == true)
            {
                Compare(view, "text after RejectAllRevisions and the Original view's", [original], [BodyText(document)]);
            }
        }
    }

    /// <summary>
    /// False when a paragraph whose mark is tracked has a complex field that does not end in it.
    /// Fields are read a paragraph at a time, so once AcceptAllRevisions or RejectAllRevisions has
    /// joined that paragraph with the next, the field goes on into the next one's text, as it
    /// does in Word, and the views cannot have the same text. Such a body is still accepted,
    /// rejected and saved; only the comparison is left out.
    /// </summary>
    private static bool JoinsKeepFields(OrdDocument document)
    {
        XNamespace main = document.Body.Name.Namespace;

        foreach (XElement paragraph in document.Body.Descendants(main + "p"))
        {
            XElement? mark = paragraph.Element(main + "pPr")?.Element(main + "rPr");

            if (mark is null || (mark.Element(main + "ins") is null && mark.Element(main + "del") is null))
            {
                continue;
            }

            int open = 0;

            // The paragraph's own fields: those in a text box inside it belong to its paragraphs.
            IEnumerable<XElement> own = paragraph.Descendants(main + "fldChar")
                .Where(fieldChar => fieldChar.Ancestors().TakeWhile(ancestor => ancestor != paragraph).Any(ancestor => ancestor.Name == main + "txbxContent") == false);

            foreach (XElement fieldChar in own)
            {
                string? type = (string?)fieldChar.Attribute(main + "fldCharType");

                if (type == "begin")
                {
                    open++;
                }
                else if (type == "end" && open > 0)
                {
                    // As the readers have it: an end with no field open ends nothing.
                    open--;
                }
            }

            if (open != 0)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// The text of every paragraph, without list labels and line or paragraph breaks.
    /// AcceptAllRevisions and RejectAllRevisions join a paragraph whose mark alone is deleted (or
    /// inserted) with the next one, as Word does, and the joined paragraph has the next one's
    /// properties, its list among them; the Accepted and Original views keep it as a paragraph of
    /// its own (docs/plan.md, known limitations). The text and its order must still be the same.
    /// </summary>
    private static string BodyText(OrdDocument document)
    {
        return string.Concat(OrdDocument.AllParagraphs(document.Blocks).Select(paragraph => paragraph.Text)).Replace("\n", string.Empty, StringComparison.Ordinal);
    }

    private static async Task<OrdDocument> OpenAsync(byte[] file, RevisionView view)
    {
        Result<OrdDocument> opened = await OrdDocument.OpenAsync(new MemoryStream(file, false), new ReadOptions { Revisions = view });

        if (opened.IsSuccess(out OrdDocument? document) == false)
        {
            throw new InvariantViolationException($"A generated body is refused: {opened.Error}");
        }

        return document;
    }

    private static void Compare(RevisionView view, string what, List<string> expected, List<string> actual)
    {
        if (expected.SequenceEqual(actual) == true)
        {
            return;
        }

        int first = 0;

        while (first < expected.Count && first < actual.Count && expected[first] == actual[first])
        {
            first++;
        }

        throw new InvariantViolationException(
            $"{view} view, {what} differ at {first} (of {expected.Count} and {actual.Count}):\n  expected:{Harness.Lines(expected.Skip(first).Take(3))}\n  actual:{Harness.Lines(actual.Skip(first).Take(3))}");
    }

}
