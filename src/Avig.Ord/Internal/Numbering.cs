using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace Avig.Ord.Internal;

/// <summary>One level of a list definition (w:lvl).</summary>
internal sealed record ListLevel(int Start, string Format, string Text);

/// <summary>
/// The list definitions of numbering.xml: w:abstractNum holds the levels, and w:num points to one
/// and may restart levels. Numbering goes on across every w:num that points to the same
/// w:abstractNum, as in Word; a w:startOverride starts it again.
/// </summary>
internal sealed class NumberingDefinitions
{
    private readonly Dictionary<string, ListLevel[]> abstracts = new(StringComparer.Ordinal);
    private readonly Dictionary<string, (string AbstractId, Dictionary<int, int> Restarts)> instances = new(StringComparer.Ordinal);

    private NumberingDefinitions()
    {
    }

    public static NumberingDefinitions Empty { get; } = new();

    public static NumberingDefinitions Read(XElement? root)
    {
        NumberingDefinitions definitions = new();

        if (root is null)
        {
            return definitions;
        }

        WordNames names = WordNames.For(root);

        foreach (XElement element in root.Elements(names.W + "abstractNum"))
        {
            string? id = element.Attribute(names.W + "abstractNumId")?.Value;

            if (id is null)
            {
                continue;
            }

            ListLevel[] levels = new ListLevel[9];

            for (int index = 0; index < levels.Length; index++)
            {
                levels[index] = new ListLevel(1, "decimal", "%" + (index + 1).ToString(CultureInfo.InvariantCulture) + ".");
            }

            foreach (XElement level in element.Elements(names.W + "lvl"))
            {
                int? number = FormatReader.Integer(level.Attribute(names.W + "ilvl")?.Value);

                if (number is null || number < 0 || number > 8)
                {
                    continue;
                }

                levels[number.Value] = new ListLevel(
                    FormatReader.Integer(names.ValueOf(level, "start")) ?? 0,
                    names.ValueOf(level, "numFmt") ?? "decimal",
                    names.ValueOf(level, "lvlText") ?? string.Empty);
            }

            definitions.abstracts[id] = levels;
        }

        foreach (XElement element in root.Elements(names.W + "num"))
        {
            string? id = element.Attribute(names.W + "numId")?.Value;
            string? abstractId = names.ValueOf(element, "abstractNumId");

            if (id is null || abstractId is null)
            {
                continue;
            }

            Dictionary<int, int> restarts = [];

            foreach (XElement levelOverride in element.Elements(names.W + "lvlOverride"))
            {
                int? level = FormatReader.Integer(levelOverride.Attribute(names.W + "ilvl")?.Value);
                int? start = FormatReader.Integer(names.ValueOf(levelOverride, "startOverride"));

                if (level is >= 0 and <= 8 && start is not null)
                {
                    restarts[level.Value] = start.Value;
                }
            }

            definitions.instances[id] = (abstractId, restarts);
        }

        return definitions;
    }

    public bool TryGet(string numberingId, out string abstractId, out ListLevel[] levels, out IReadOnlyDictionary<int, int> restarts)
    {
        abstractId = string.Empty;
        levels = [];
        restarts = new Dictionary<int, int>();

        if (instances.TryGetValue(numberingId, out (string AbstractId, Dictionary<int, int> Restarts) instance) == false
            || abstracts.TryGetValue(instance.AbstractId, out ListLevel[]? found) == false)
        {
            return false;
        }

        abstractId = instance.AbstractId;
        levels = found;
        restarts = instance.Restarts;

        return true;
    }
}

/// <summary>
/// Counts list items in document order and makes their labels, as Word shows them: "1.",
/// "2.3", "a)", "iv." or a bullet.
/// </summary>
internal sealed class ListCounter
{
    /// <summary>The largest number written in letters: Word takes no start above 32 767.</summary>
    private const int MaxLetterNumber = 32_767;

    private readonly NumberingDefinitions definitions;
    private readonly Dictionary<string, int[]> counters = new(StringComparer.Ordinal);
    private readonly HashSet<string> restarted = new(StringComparer.Ordinal);

    public ListCounter(NumberingDefinitions definitions)
    {
        this.definitions = definitions;
    }

    /// <summary>The next item of a list at a level; false when the numbering id is 0 (no list) or unknown.</summary>
    public bool TryNext(string numberingId, int level, out ListInfo info)
    {
        info = new ListInfo(0, string.Empty, false, numberingId);

        if (numberingId == "0" || definitions.TryGet(numberingId, out string abstractId, out ListLevel[] levels, out IReadOnlyDictionary<int, int> restarts) == false)
        {
            return false;
        }

        int clamped = Math.Clamp(level, 0, 8);

        if (counters.TryGetValue(abstractId, out int[]? counts) == false)
        {
            counts = new int[9];

            for (int index = 0; index < counts.Length; index++)
            {
                counts[index] = levels[index].Start - 1;
            }

            counters[abstractId] = counts;
        }

        // A w:num with a startOverride starts its levels again the first time it is used.
        if (restarts.Count > 0 && restarted.Add(numberingId) == true)
        {
            foreach ((int restartLevel, int start) in restarts)
            {
                counts[restartLevel] = start - 1;
            }
        }

        counts[clamped]++;

        for (int deeper = clamped + 1; deeper < counts.Length; deeper++)
        {
            counts[deeper] = levels[deeper].Start - 1;
        }

        ListLevel definition = levels[clamped];
        bool bullet = definition.Format == "bullet";
        string label = Label(definition, levels, counts, bullet);
        info = new ListInfo(clamped, label, bullet, numberingId) { Number = counts[clamped], NumberFormat = definition.Format };

        return true;
    }

    private static string Label(ListLevel definition, ListLevel[] levels, int[] counts, bool bullet)
    {
        if (bullet == true)
        {
            return Bullet(definition.Text);
        }

        StringBuilder label = new();
        string text = definition.Text;

        for (int index = 0; index < text.Length; index++)
        {
            if (text[index] == '%' && index + 1 < text.Length && text[index + 1] is >= '1' and <= '9')
            {
                int referenced = text[index + 1] - '1';
                label.Append(FormatNumber(Math.Max(counts[referenced], 0), levels[referenced].Format));
                index++;

                continue;
            }

            label.Append(text[index]);
        }

        return label.ToString();
    }

    /// <summary>A bullet as it reads without Word's symbol fonts: the private-use characters Symbol and Wingdings bullets are written with become their Unicode look-alikes.</summary>
    private static string Bullet(string text)
    {
        return text switch
        {
            "" or "·" or "" => "•",
            "" => "▪",
            "" => "➢",
            "" => "✓",
            "o" => "◦",
            _ => text,
        };
    }

    private static string FormatNumber(int number, string format)
    {
        return format switch
        {
            "lowerLetter" => Letters(number, 'a'),
            "upperLetter" => Letters(number, 'A'),
            "lowerRoman" => Roman(number).ToLowerInvariant(),
            "upperRoman" => Roman(number),
            "decimalZero" => number.ToString("00", CultureInfo.InvariantCulture),
            "none" => string.Empty,
            "ordinal" => number.ToString(CultureInfo.InvariantCulture) + ".",
            _ => number.ToString(CultureInfo.InvariantCulture),
        };
    }

    /// <summary>a … z, then aa … zz, as Word counts letters. Past 32 767, the largest start Word takes, the number in digits.</summary>
    private static string Letters(int number, char first)
    {
        if (number <= 0)
        {
            return string.Empty;
        }

        // A start of int.MaxValue in a file made an 82-million-letter label for every item.
        if (number > MaxLetterNumber)
        {
            return number.ToString(CultureInfo.InvariantCulture);
        }

        char letter = (char)(first + ((number - 1) % 26));

        return new string(letter, ((number - 1) / 26) + 1);
    }

    private static string Roman(int number)
    {
        if (number is <= 0 or >= 4000)
        {
            return number.ToString(CultureInfo.InvariantCulture);
        }

        (int Value, string Symbol)[] numerals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")];
        StringBuilder roman = new();
        int rest = number;

        foreach ((int value, string symbol) in numerals)
        {
            while (rest >= value)
            {
                roman.Append(symbol);
                rest -= value;
            }
        }

        return roman.ToString();
    }
}
