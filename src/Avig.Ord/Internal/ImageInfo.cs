using System.Buffers.Binary;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Avig.Ord.Internal;

/// <summary>A picture's size in pixels, read from its header, without an image library.</summary>
internal static partial class ImageInfo
{
    /// <summary>The size in pixels of a PNG, JPEG, GIF, BMP or SVG (from width/height or viewBox); false when it cannot be told.</summary>
    public static bool TryGetSize(ReadOnlySpan<byte> data, out int width, out int height)
    {
        width = 0;
        height = 0;

        if (data.Length >= 24 && data.StartsWith((ReadOnlySpan<byte>)[0x89, 0x50, 0x4E, 0x47]) == true)
        {
            width = BinaryPrimitives.ReadInt32BigEndian(data[16..]);
            height = BinaryPrimitives.ReadInt32BigEndian(data[20..]);
        }
        else if (data.Length >= 10 && data.StartsWith("GIF8"u8) == true)
        {
            width = BinaryPrimitives.ReadUInt16LittleEndian(data[6..]);
            height = BinaryPrimitives.ReadUInt16LittleEndian(data[8..]);
        }
        else if (data.Length >= 26 && data.StartsWith("BM"u8) == true)
        {
            width = BinaryPrimitives.ReadInt32LittleEndian(data[18..]);
            height = Math.Abs(BinaryPrimitives.ReadInt32LittleEndian(data[22..]));
        }
        else if (data.Length >= 4 && data[0] == 0xFF && data[1] == 0xD8)
        {
            Jpeg(data, out width, out height);
        }
        else
        {
            Svg(data, out width, out height);
        }

        return width > 0 && height > 0 && width < 1_000_000 && height < 1_000_000;
    }

    /// <summary>Walks the JPEG segments to the first start-of-frame, which has the size.</summary>
    private static void Jpeg(ReadOnlySpan<byte> data, out int width, out int height)
    {
        width = 0;
        height = 0;
        int position = 2;

        while (position + 9 < data.Length)
        {
            if (data[position] != 0xFF)
            {
                position++;

                continue;
            }

            byte marker = data[position + 1];
            int length = BinaryPrimitives.ReadUInt16BigEndian(data[(position + 2)..]);

            if (marker is >= 0xC0 and <= 0xCF and not 0xC4 and not 0xC8 and not 0xCC)
            {
                height = BinaryPrimitives.ReadUInt16BigEndian(data[(position + 5)..]);
                width = BinaryPrimitives.ReadUInt16BigEndian(data[(position + 7)..]);

                return;
            }

            if (length < 2)
            {
                return;
            }

            position += 2 + length;
        }
    }

    private static void Svg(ReadOnlySpan<byte> data, out int width, out int height)
    {
        width = 0;
        height = 0;
        string head = System.Text.Encoding.UTF8.GetString(data[..Math.Min(data.Length, 4096)]);
        Match root = SvgRoot().Match(head);

        if (root.Success == false)
        {
            return;
        }

        string attributes = root.Groups[1].Value;
        Match widthMatch = SvgWidth().Match(attributes);
        Match heightMatch = SvgHeight().Match(attributes);

        if (widthMatch.Success == true && heightMatch.Success == true
            && double.TryParse(widthMatch.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double svgWidth) == true
            && double.TryParse(heightMatch.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double svgHeight) == true)
        {
            width = (int)Math.Round(svgWidth);
            height = (int)Math.Round(svgHeight);

            return;
        }

        Match box = SvgViewBox().Match(attributes);

        if (box.Success == true
            && double.TryParse(box.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double boxWidth) == true
            && double.TryParse(box.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double boxHeight) == true)
        {
            width = (int)Math.Round(boxWidth);
            height = (int)Math.Round(boxHeight);
        }
    }

    [GeneratedRegex(@"<svg\b([^>]*)>", RegexOptions.IgnoreCase, 1000)]
    private static partial Regex SvgRoot();

    [GeneratedRegex(@"\bwidth\s*=\s*[""']\s*([0-9.]+)(px)?\s*[""']", RegexOptions.IgnoreCase, 1000)]
    private static partial Regex SvgWidth();

    [GeneratedRegex(@"\bheight\s*=\s*[""']\s*([0-9.]+)(px)?\s*[""']", RegexOptions.IgnoreCase, 1000)]
    private static partial Regex SvgHeight();

    [GeneratedRegex(@"\bviewBox\s*=\s*[""']\s*[-0-9.]+[\s,]+[-0-9.]+[\s,]+([0-9.]+)[\s,]+([0-9.]+)\s*[""']", RegexOptions.IgnoreCase, 1000)]
    private static partial Regex SvgViewBox();
}
