using System.Xml.Linq;

namespace Avig.Ord.Internal;

/// <summary>Combines run properties (w:rPr) in the order the schema wants them.</summary>
internal static class RunProperties
{
    /// <summary>The order of w:rPr's children in the schema (CT_RPr).</summary>
    private static readonly string[] Order =
    [
        "ins", "del", "moveFrom", "moveTo",
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
        "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u",
        "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange",
    ];

    /// <summary>The properties of <paramref name="over"/> where it has them, those of <paramref name="under"/> where it does not, sorted.</summary>
    public static XElement? Merge(WordNames names, XElement? under, XElement? over)
    {
        if (over is null)
        {
            return under;
        }

        Dictionary<string, XElement> merged = new(StringComparer.Ordinal);

        foreach (XElement element in under?.Elements() ?? [])
        {
            merged[element.Name.LocalName] = new XElement(element);
        }

        foreach (XElement element in over.Elements())
        {
            merged[element.Name.LocalName] = new XElement(element);
        }

        return new XElement(names.RunProperties, merged.Values.OrderBy(element => Rank(element.Name.LocalName)));
    }

    private static int Rank(string localName)
    {
        int index = Array.IndexOf(Order, localName);

        if (index < 0)
        {
            return Order.Length - 1;
        }

        return index;
    }
}
