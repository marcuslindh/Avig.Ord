using System.Text;

namespace Avig.Ord.Fuzz;

/// <summary>
/// <c>flat-opc</c> — the every-feature document as <see cref="DocxWriter"/> writes it in Flat
/// OPC (one XML file, "Word XML Document"), mutated as text. This reaches the Flat OPC parser of
/// Avig.OpenXml (part names, content types, base64 binary data, duplicate parts) and, through it,
/// the same parts as <c>docx-part</c>.
/// </summary>
internal static class FlatTarget
{
    private static readonly string Flat = Encoding.UTF8.GetString(Template.Write(true));

    private static readonly string[] FlatTokens =
    [
        "<pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/xml\"><pkg:xmlData><x/></pkg:xmlData></pkg:part>",
        "<pkg:part pkg:name=\"/../../x.xml\" pkg:contentType=\"application/xml\"><pkg:xmlData><x/></pkg:xmlData></pkg:part>",
        "<pkg:part pkg:name=\"\" pkg:contentType=\"\"/>", "<pkg:part>", "</pkg:part>", "<pkg:binaryData>", "</pkg:binaryData>", "<pkg:binaryData>!!!not base64</pkg:binaryData>",
        "<pkg:binaryData>" + new string('A', 30_000) + "</pkg:binaryData>", " pkg:compression=\"store\"", " pkg:padding=\"512\"", " pkg:name=\"/word/styles.xml\"",
        " pkg:contentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"", "<pkg:xmlData>", "</pkg:xmlData>", "<?mso-application progid=\"Excel.Sheet\"?>",
    ];

    private static readonly string[] Tokens = [.. PartTarget.Tokens, .. FlatTokens];

    public static FuzzTarget Create()
    {
        return new FuzzTarget("flat-opc", "a Flat OPC document mutated as text; read with both readers, change, save, reopen", RunAsync, Generate);
    }

    private static byte[] Generate(Random random)
    {
        string content = Flat;

        if (random.Next(10) > 0)
        {
            content = Inputs.Mutate(random, content, Tokens);
        }

        byte[] text = Encoding.UTF8.GetBytes(content);
        byte[] input = new byte[text.Length + 1];
        input[0] = (byte)random.Next(256);
        text.CopyTo(input, 1);

        return input;
    }

    private static async Task RunAsync(byte[] input)
    {
        if (input.Length < 1)
        {
            return;
        }

        await Harness.RunFileAsync(input[1..], input[0]);
    }
}
