using System.Text;

namespace Avig.Ord.Fuzz;

/// <summary>
/// Reads values from fuzz input. libFuzzer mutates bytes, so every target that needs more
/// than a string turns its bytes into values here, and a short input simply runs out into zeros.
/// </summary>
internal sealed class ByteSource
{
    private readonly byte[] data;
    private int position;

    public ByteSource(byte[] data)
    {
        this.data = data;
    }

    public bool HasMore => position < data.Length;

    public byte NextByte()
    {
        if (position >= data.Length)
        {
            return 0;
        }

        byte value = data[position];
        position++;

        return value;
    }

    public long NextLong()
    {
        long value = 0;

        for (int index = 0; index < 8; index++)
        {
            value = (value << 8) | NextByte();
        }

        return value;
    }

    public double NextDouble()
    {
        return BitConverter.Int64BitsToDouble(NextLong());
    }

    public char NextChar()
    {
        return (char)((NextByte() << 8) | NextByte());
    }
}

/// <summary>Building blocks for the smoke-mode generators.</summary>
internal static class Inputs
{
    /// <summary>
    /// Characters that have caused trouble in Office XML code somewhere: XML-invalid controls, lone
    /// surrogates, the noncharacters, the _xHHHH_ escape of SpreadsheetML, line breaks (XML folds
    /// \r\n into \n on read), leading and trailing spaces (lost without xml:space).
    /// </summary>
    private static readonly string[] Awkward =
    [
        "\0", "\u0001", "\u0008", "\t", "\n", "\r", "\r\n", "\u000B", "\u001F", " ", "  ",
        "\uD800", "\uDC00", "😀", "￾", "￿", " ", "​", "\u0085",
        "_x0041_", "_x005F_", "_x", "_", "_xD800_", "_x0000_", "_X0041_", "_x00411_", "_xc1e5", "_x00e9_", "_x0041",
        "&", "<", ">", "\"", "'", "&amp;", "]]>", "Öst", "ﬁ", "İ",
        "=", "+", "-", "@", "1", "1,5", "1e5", "TRUE", "#N/A",
    ];

    public static T Pick<T>(Random random, IReadOnlyList<T> items)
    {
        return items[random.Next(items.Count)];
    }

    /// <summary>Text built mostly from ordinary letters with awkward pieces mixed in.</summary>
    public static string Text(Random random, int maxPieces)
    {
        StringBuilder builder = new();
        int pieces = random.Next(maxPieces + 1);

        for (int piece = 0; piece < pieces; piece++)
        {
            if (random.Next(3) == 0)
            {
                builder.Append(Pick(random, Awkward));
            }
            else
            {
                builder.Append((char)('a' + random.Next(26)));
            }
        }

        return builder.ToString();
    }

    /// <summary>
    /// A few textual mutations of <paramref name="text"/>: cut a span, repeat a span, put in a
    /// token, or swap a number for an edge value. Close enough to the original that the parser
    /// gets past its first checks, which random bytes never do.
    /// </summary>
    public static string Mutate(Random random, string text, IReadOnlyList<string> tokens)
    {
        StringBuilder builder = new(text);
        int mutations = random.Next(1, 5);

        for (int mutation = 0; mutation < mutations; mutation++)
        {
            int at = random.Next(builder.Length + 1);
            int kind = random.Next(10);

            if (kind < 2 && builder.Length > 0)
            {
                int length = Math.Min(random.Next(1, 24), builder.Length - Math.Min(at, builder.Length - 1));
                builder.Remove(Math.Min(at, builder.Length - 1), length);
            }
            else if (kind < 3 && builder.Length > 0)
            {
                int start = random.Next(builder.Length);
                int length = Math.Min(random.Next(1, 64), builder.Length - start);
                builder.Insert(at, builder.ToString(start, length));
            }
            else if (kind < 7)
            {
                builder.Insert(at, Pick(random, tokens));
            }
            else
            {
                ReplaceNumber(random, builder);
            }
        }

        return builder.ToString();
    }

    private static readonly string[] EdgeNumbers =
    [
        "0", "-1", "1", "2", "9", "10", "26", "27", "255", "256", "65535", "65536", "1048576",
        "1048577", "16384", "16385", "2147483647", "2147483648", "-2147483648", "99999999999999999999",
        "1e308", "1e309", "-0", "0.5", "1E-320", "NaN", "",
    ];

    private static void ReplaceNumber(Random random, StringBuilder builder)
    {
        int start = random.Next(builder.Length + 1);

        while (start < builder.Length && char.IsAsciiDigit(builder[start]) == false)
        {
            start++;
        }

        int end = start;

        while (end < builder.Length && char.IsAsciiDigit(builder[end]) == true)
        {
            end++;
        }

        builder.Remove(start, end - start);
        builder.Insert(start, Pick(random, EdgeNumbers));
    }

    /// <summary>A string as UTF-16 bytes, so that a target can be handed lone surrogates, which UTF-8 cannot carry.</summary>
    public static byte[] Utf16(string text)
    {
        // By hand: Encoding.BigEndianUnicode would replace a lone surrogate with U+FFFD.
        byte[] bytes = new byte[text.Length * 2];

        for (int index = 0; index < text.Length; index++)
        {
            bytes[index * 2] = (byte)(text[index] >> 8);
            bytes[(index * 2) + 1] = (byte)text[index];
        }

        return bytes;
    }

    /// <summary>The reverse of <see cref="Utf16"/>, keeping lone surrogates as they are.</summary>
    public static string FromUtf16(ReadOnlySpan<byte> data)
    {
        char[] chars = new char[data.Length / 2];

        for (int index = 0; index < chars.Length; index++)
        {
            chars[index] = (char)((data[index * 2] << 8) | data[(index * 2) + 1]);
        }

        return new string(chars);
    }
}
