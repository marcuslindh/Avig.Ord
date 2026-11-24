namespace Avig.Ord.Fuzz;

/// <summary>
/// <c>zip-raw</c> — a damaged zip container must give a Result failure with an error kind,
/// never an exception.
///
/// <c>docx-part</c> keeps the zip valid and mutates the XML inside it. This target does the
/// opposite: it mutates the bytes of the valid rich template document, where the zip structure is.
/// That covers local headers and the central directory (sizes, offsets, compression method,
/// name lengths, flags), a truncated file, and junk in front of or behind the archive.
/// </summary>
internal static class ZipTarget
{
    private static readonly byte[] Template = Fuzz.Template.Build(Fuzz.Template.Rich, string.Empty, []).ToArray();

    /// <summary>The fields worth hitting, as offsets from a local (PK\x03\x04) or central (PK\x01\x02) header.</summary>
    private static readonly int[] LocalFields = [4, 6, 8, 14, 18, 22, 26, 28];

    private static readonly int[] CentralFields = [6, 8, 10, 16, 20, 24, 28, 30, 32, 34, 38, 42];

    private static readonly uint[] EdgeValues = [0, 1, 0x7F, 0xFF, 0xFFFF, 0xFFFF_FFFF, 0x7FFF_FFFF, 0x8000_0000, 14, 12, 99, 9];

    public static FuzzTarget Create()
    {
        return new FuzzTarget("zip-raw", "the bytes of a valid document mutated where the zip structure is", RunAsync, Generate);
    }

    private static byte[] Generate(Random random)
    {
        List<byte> bytes = [.. Template];
        int mutations = random.Next(1, 4);

        for (int mutation = 0; mutation < mutations; mutation++)
        {
            int kind = random.Next(10);

            if (kind < 6)
            {
                HitHeaderField(random, bytes);
            }
            else if (kind == 6)
            {
                int keep = random.Next(bytes.Count / 2, bytes.Count);
                bytes.RemoveRange(keep, bytes.Count - keep);
            }
            else if (kind == 7)
            {
                bytes.InsertRange(0, Enumerable.Repeat((byte)random.Next(256), random.Next(1, 64)));
            }
            else if (kind == 8)
            {
                int at = random.Next(bytes.Count);
                bytes[at] = (byte)random.Next(256);
            }
            else
            {
                int start = random.Next(bytes.Count);
                int length = Math.Min(random.Next(1, 200), bytes.Count - start);
                bytes.InsertRange(random.Next(bytes.Count), bytes.GetRange(start, length));
            }

            if (bytes.Count == 0)
            {
                bytes.AddRange(Template);
            }
        }

        return [.. bytes];
    }

    /// <summary>Writes an edge value into a field of a randomly chosen local or central header.</summary>
    private static void HitHeaderField(Random random, List<byte> bytes)
    {
        List<(int Offset, bool Central)> headers = [];

        for (int position = 0; position < bytes.Count - 4; position++)
        {
            if (bytes[position] == 0x50 && bytes[position + 1] == 0x4B && bytes[position + 2] is 1 or 3 && bytes[position + 3] == bytes[position + 2] + 1)
            {
                headers.Add((position, bytes[position + 2] == 1));
            }
        }

        if (headers.Count == 0)
        {
            return;
        }

        (int offset, bool central) = headers[random.Next(headers.Count)];
        int[] fields = LocalFields;

        if (central == true)
        {
            fields = CentralFields;
        }

        int field = offset + fields[random.Next(fields.Length)];
        uint value = EdgeValues[random.Next(EdgeValues.Length)];
        int width = 4;

        if (random.Next(2) == 0)
        {
            width = 2;
        }

        for (int index = 0; index < width && field + index < bytes.Count; index++)
        {
            bytes[field + index] = (byte)(value >> (8 * index));
        }
    }

    private static async Task RunAsync(byte[] input)
    {
        // The view is fixed: the bytes of the input are the file, with nothing to spare for a selector.
        await Harness.RunFileAsync(input, 0);
    }
}
