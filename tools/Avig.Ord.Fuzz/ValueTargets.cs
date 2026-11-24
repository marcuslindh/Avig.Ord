using System.Buffers.Binary;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Avig.Ord.Internal;

namespace Avig.Ord.Fuzz;

/// <summary>Targets for the small parsers: picture headers and the formatting in w:rPr and w:pPr.</summary>
internal static class ValueTargets
{
    private static readonly string[] EdgeValues =
    [
        "0", "1", "-1", "20", "240", "2147483647", "2147483648", "-2147483648", "9223372036854775807", "-9223372036854775808", "99999999999999999999",
        "14523000000000000", "20000000000000000", "-20000000000000000", "1.5", "1e309", "NaN", "", " 1", "12pt", "1.5cm", "-3mm", "2in", "1pc", "3pi", "1e308pt", "NaNpt", "Infinitypt", "-1e309cm", "pt", "x",
        "true", "false", "on", "off", "none", "single", "superscript", "center", "both", "start", "end",
    ];

    private static readonly string[] RunElements = ["b", "i", "u", "strike", "dstrike", "sz", "rFonts", "color", "highlight", "vertAlign", "caps", "smallCaps", "vanish"];

    private static readonly string[] ParagraphElements = ["jc", "spacing", "ind", "outlineLvl", "keepNext", "pageBreakBefore"];

    private static readonly string[] Attributes = ["val", "before", "after", "left", "right", "start", "end", "firstLine", "hanging", "ascii"];

    /// <summary><c>image-size</c> — the header of a PNG, GIF, BMP, JPEG or SVG, mutated: no exception, and a size it gives is the one in the header.</summary>
    public static FuzzTarget ImageSize()
    {
        return new FuzzTarget("image-size", "ImageInfo.TryGetSize on picture headers: no exception, and the size in the header", RunImageAsync, GenerateImage);
    }

    /// <summary><c>format</c> — w:rPr and w:pPr with edge values: no exception, and lengths that do not wrap around.</summary>
    public static FuzzTarget Format()
    {
        return new FuzzTarget("format", "FormatReader on w:rPr and w:pPr with edge values: no exception, lengths that do not overflow", RunFormatAsync, GenerateFormat);
    }

    private static byte[] GenerateImage(Random random)
    {
        int kind = random.Next(5);
        List<byte> data = [];
        uint width = (uint)EdgeNumber(random);
        uint height = (uint)EdgeNumber(random);

        switch (kind)
        {
            case 0:
                data.AddRange([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, (byte)'I', (byte)'H', (byte)'D', (byte)'R']);
                data.AddRange(BigEndian(width));
                data.AddRange(BigEndian(height));
                data.AddRange([8, 6, 0, 0, 0]);

                break;

            case 1:
                data.AddRange("GIF89a"u8.ToArray());
                data.AddRange([(byte)width, (byte)(width >> 8), (byte)height, (byte)(height >> 8), 0, 0, 0]);

                break;

            case 2:
                data.AddRange("BM"u8.ToArray());
                data.AddRange(new byte[16]);
                data.AddRange(LittleEndian(width));
                data.AddRange(LittleEndian(height));
                data.AddRange([1, 0, 24, 0]);

                break;

            case 3:
                data.AddRange([0xFF, 0xD8]);
                int segments = random.Next(0, 6);

                for (int segment = 0; segment < segments; segment++)
                {
                    int length = random.Next(0, 40);
                    data.AddRange([0xFF, (byte)random.Next(0xC0, 0xFF), (byte)(length >> 8), (byte)length]);
                    data.AddRange(Enumerable.Range(0, Math.Max(length - 2, 0)).Select(_ => (byte)random.Next(256)));
                }

                data.AddRange([0xFF, 0xC0, 0, 17, 8, (byte)(height >> 8), (byte)height, (byte)(width >> 8), (byte)width, 3]);

                break;

            default:
                string[] sizes = ["40", "1e5", "-3", "0", "99999999999", "1.5", "NaN", "", "...", "40px", "40%", "1,5"];
                string svg = $"<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{Inputs.Pick(random, sizes)}\" height=\"{Inputs.Pick(random, sizes)}\" viewBox=\"0 0 {Inputs.Pick(random, sizes)} {Inputs.Pick(random, sizes)}\">";

                if (random.Next(4) == 0)
                {
                    svg = new string(' ', random.Next(4_000, 5_000)) + svg;
                }

                data.AddRange(Encoding.UTF8.GetBytes(svg));

                break;
        }

        // Cut or flip a byte now and then: a header that ends early is the classic out-of-range read.
        if (random.Next(3) == 0 && data.Count > 0)
        {
            int keep = random.Next(data.Count);
            data.RemoveRange(keep, data.Count - keep);
        }

        if (random.Next(3) == 0 && data.Count > 0)
        {
            data[random.Next(data.Count)] = (byte)random.Next(256);
        }

        return [.. data];
    }

    private static Task RunImageAsync(byte[] input)
    {
        ReadOnlySpan<byte> data = input;
        bool found = ImageInfo.TryGetSize(data, out int width, out int height);
        _ = WordTypes.ImageContentType(data, out _);

        if (found == true && (width <= 0 || height <= 0 || width >= 1_000_000 || height >= 1_000_000))
        {
            throw new InvariantViolationException($"TryGetSize gave {width}x{height}.");
        }

        // A PNG's size is in its header, at a fixed place: what is there and in range must be found.
        if (data.Length >= 24 && data.StartsWith((ReadOnlySpan<byte>)[0x89, 0x50, 0x4E, 0x47]) == true)
        {
            int pngWidth = BinaryPrimitives.ReadInt32BigEndian(data[16..]);
            int pngHeight = BinaryPrimitives.ReadInt32BigEndian(data[20..]);

            if (pngWidth is > 0 and < 1_000_000 && pngHeight is > 0 and < 1_000_000 && (found == false || width != pngWidth || height != pngHeight))
            {
                throw new InvariantViolationException($"A PNG of {pngWidth}x{pngHeight} gave {found}, {width}x{height}.");
            }
        }

        return Task.CompletedTask;
    }

    private static byte[] GenerateFormat(Random random)
    {
        StringBuilder xml = new();
        bool run = random.Next(2) == 0;
        string[] elements = ParagraphElements;
        string root = "pPr";

        if (run == true)
        {
            elements = RunElements;
            root = "rPr";
        }

        xml.Append("<w:").Append(root).Append(" xmlns:w=\"").Append(Inputs.Pick(random, [WordNames.TransitionalMain, WordNames.StrictMain])).Append("\">");
        int count = random.Next(1, 8);

        for (int index = 0; index < count; index++)
        {
            xml.Append("<w:").Append(Inputs.Pick(random, elements));
            int attributes = random.Next(0, 4);

            for (int attribute = 0; attribute < attributes; attribute++)
            {
                xml.Append(" w:").Append(Inputs.Pick(random, Attributes)).Append("=\"").Append(Inputs.Pick(random, EdgeValues)).Append('"');
            }

            xml.Append("/>");
        }

        xml.Append("</w:").Append(root).Append('>');

        return Encoding.UTF8.GetBytes(xml.ToString());
    }

    private static Task RunFormatAsync(byte[] input)
    {
        XElement element;

        try
        {
            element = XElement.Parse(Encoding.UTF8.GetString(input));
        }
        catch (System.Xml.XmlException)
        {
            // Not XML: the package never gets there, OpcXml refuses it first.
            return Task.CompletedTask;
        }

        WordNames names = WordNames.For(element);
        _ = FormatReader.Run(names, element);
        _ = FormatReader.Paragraph(names, element);

        foreach (XAttribute attribute in element.Descendants().Attributes())
        {
            CheckTwips(attribute.Value);
            _ = FormatReader.Emu(attribute.Value);
            _ = FormatReader.Integer(attribute.Value);
        }

        return Task.CompletedTask;
    }

    /// <summary>A whole number of twips gives that many twips, or nothing when it is too large to hold: never a length of the other sign.</summary>
    private static void CheckTwips(string value)
    {
        Length? length = FormatReader.Twips(value);

        if (length is null || long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long twips) == false)
        {
            return;
        }

        if (Math.Abs((double)twips) < 1e12 && length.Value.Emu != twips * Length.EmuPerTwip)
        {
            throw new InvariantViolationException($"'{value}' twips gave {length.Value.Emu} EMU.");
        }

        if (twips != 0 && Math.Sign(twips) != Math.Sign(length.Value.Emu))
        {
            throw new InvariantViolationException($"'{value}' twips gave {length.Value.Emu} EMU, of the other sign: the multiplication overflowed.");
        }
    }

    private static long EdgeNumber(Random random)
    {
        long[] numbers = [0, 1, 2, 40, 255, 256, 65_535, 65_536, 999_999, 1_000_000, int.MaxValue, 0x8000_0000, 0xFFFF_FFFF];

        return numbers[random.Next(numbers.Length)];
    }

    private static byte[] BigEndian(uint value)
    {
        byte[] bytes = new byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(bytes, value);

        return bytes;
    }

    private static byte[] LittleEndian(uint value)
    {
        byte[] bytes = new byte[4];
        BinaryPrimitives.WriteUInt32LittleEndian(bytes, value);

        return bytes;
    }
}
