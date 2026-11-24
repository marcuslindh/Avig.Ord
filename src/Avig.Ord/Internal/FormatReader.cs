using System.Globalization;
using System.Xml.Linq;

namespace Avig.Ord.Internal;

/// <summary>Reads w:rPr and w:pPr into <see cref="RunFormat"/> and <see cref="ParagraphFormat"/>.</summary>
internal static class FormatReader
{
    public static RunFormat Run(WordNames names, XElement? properties)
    {
        if (properties is null)
        {
            return RunFormat.None;
        }

        return new RunFormat
        {
            Bold = names.ToggleOf(properties, "b"),
            Italic = names.ToggleOf(properties, "i"),
            Underline = Underline(names, properties),
            Strike = Strike(names, properties),
            Size = HalfPoints(names.ValueOf(properties, "sz")),
            Font = properties.Element(names.W + "rFonts")?.Attribute(names.W + "ascii")?.Value,
            Color = names.ValueOf(properties, "color"),
            Highlight = names.ValueOf(properties, "highlight"),
            Position = Position(names.ValueOf(properties, "vertAlign")),
            AllCaps = names.ToggleOf(properties, "caps"),
            SmallCaps = names.ToggleOf(properties, "smallCaps"),
            Hidden = names.ToggleOf(properties, "vanish"),
        };
    }

    public static ParagraphFormat Paragraph(WordNames names, XElement? properties)
    {
        if (properties is null)
        {
            return ParagraphFormat.None;
        }

        XElement? spacing = properties.Element(names.W + "spacing");
        XElement? indent = properties.Element(names.W + "ind");

        return new ParagraphFormat
        {
            Alignment = Align(names.ValueOf(properties, "jc")),
            SpaceBefore = Twips(spacing?.Attribute(names.W + "before")?.Value),
            SpaceAfter = Twips(spacing?.Attribute(names.W + "after")?.Value),
            IndentStart = Twips(indent?.Attribute(names.W + "start")?.Value ?? indent?.Attribute(names.W + "left")?.Value),
            IndentEnd = Twips(indent?.Attribute(names.W + "end")?.Value ?? indent?.Attribute(names.W + "right")?.Value),
            FirstLineIndent = FirstLine(names, indent),
            OutlineLevel = Outline(names.ValueOf(properties, "outlineLvl")),
            KeepWithNext = names.ToggleOf(properties, "keepNext"),
            PageBreakBefore = names.ToggleOf(properties, "pageBreakBefore"),
        };
    }

    /// <summary>A whole number in a w:val, or null when there is none or it is not a number.</summary>
    public static int? Integer(string? value)
    {
        int? number = null;

        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) == true)
        {
            number = parsed;
        }

        return number;
    }

    /// <summary>A length in twips (also written as "12pt" or "1.5cm" in Strict and in some generators).</summary>
    public static Length? Twips(string? value)
    {
        Length? length = null;

        if (value is not null)
        {
            if (long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long twips) == true)
            {
                length = Length.FromTwips(twips);
            }
            else if (TryUniversal(value, out Length universal) == true)
            {
                length = universal;
            }
        }

        return length;
    }

    /// <summary>A length in EMUs, as on a picture's wp:extent.</summary>
    public static Length? Emu(string? value)
    {
        Length? length = null;

        if (long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu) == true)
        {
            length = Length.FromEmu(emu);
        }

        return length;
    }

    private static Length? HalfPoints(string? value)
    {
        Length? length = null;

        if (long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long halfPoints) == true)
        {
            length = Length.FromHalfPoints(halfPoints);
        }
        else if (value is not null && TryUniversal(value, out Length universal) == true)
        {
            length = universal;
        }

        return length;
    }

    /// <summary>A "universal measure" (ST_UniversalMeasure): a number followed by mm, cm, in, pt, pc or pi.</summary>
    private static bool TryUniversal(string value, out Length length)
    {
        length = Length.Zero;

        if (value.Length < 3)
        {
            return false;
        }

        string unit = value[^2..];

        if (double.TryParse(value[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out double number) == false)
        {
            return false;
        }

        switch (unit)
        {
            case "mm":
                length = Length.Mm(number);

                return true;

            case "cm":
                length = Length.Cm(number);

                return true;

            case "in":
                length = Length.Inch(number);

                return true;

            case "pt":
                length = Length.Pt(number);

                return true;

            case "pc":
            case "pi":
                length = Length.Pt(number * 12);

                return true;

            default:
                return false;
        }
    }

    private static bool? Underline(WordNames names, XElement properties)
    {
        XElement? underline = properties.Element(names.W + "u");
        bool? state = null;

        if (underline is not null)
        {
            state = underline.Attribute(names.Val)?.Value is not "none";
        }

        return state;
    }

    private static bool? Strike(WordNames names, XElement properties)
    {
        bool? single = names.ToggleOf(properties, "strike");
        bool? both = names.ToggleOf(properties, "dstrike");

        if (single is null)
        {
            return both;
        }

        return single == true || both == true;
    }

    private static VerticalPosition? Position(string? value)
    {
        VerticalPosition? position = value switch
        {
            "superscript" => VerticalPosition.Superscript,
            "subscript" => VerticalPosition.Subscript,
            "baseline" => VerticalPosition.Baseline,
            _ => default(VerticalPosition?),
        };

        return position;
    }

    private static Alignment? Align(string? value)
    {
        Alignment? alignment = value switch
        {
            "left" or "start" => Ord.Alignment.Start,
            "center" => Ord.Alignment.Center,
            "right" or "end" => Ord.Alignment.End,
            "both" or "distribute" or "lowKashida" or "mediumKashida" or "highKashida" or "thaiDistribute" => Ord.Alignment.Justified,
            _ => default(Alignment?),
        };

        return alignment;
    }

    private static Length? FirstLine(WordNames names, XElement? indent)
    {
        Length? first = Twips(indent?.Attribute(names.W + "firstLine")?.Value);
        Length? hanging = Twips(indent?.Attribute(names.W + "hanging")?.Value);

        if (hanging is not null)
        {
            return Length.Zero - hanging.Value;
        }

        return first;
    }

    /// <summary>The outline level: 0–8 a heading level, 9 body text (kept, so that it can override a heading style).</summary>
    private static int? Outline(string? value)
    {
        int? level = Integer(value);

        if (level is < 0 or > 9)
        {
            level = null;
        }

        return level;
    }
}
