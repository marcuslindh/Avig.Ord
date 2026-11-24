using System.Collections;
using System.Globalization;

namespace Avig.Ord;

/// <summary>
/// Values for a template, by name. A class marked <see cref="OrdTemplateAttribute"/> gets this
/// made by the source generator; <see cref="TemplateValues"/> holds values in a dictionary.
/// </summary>
public interface ITemplateData
{
    /// <summary>The value for a placeholder or section name; false when there is none.</summary>
    bool TryGetValue(string name, out object? value);
}

/// <summary>
/// Marks a partial class as data for a template: the source generator makes it an
/// <see cref="ITemplateData"/> whose names are its public properties, without reflection, so
/// it works with trimming and Native AOT. A property that is a list of such classes is a
/// section that repeats.
/// </summary>
/// <example>
/// <code>
/// [OrdTemplate]
/// public partial class Invoice
/// {
///     public string Customer { get; init; } = "";
///     public DateOnly Date { get; init; }
///     public IReadOnlyList&lt;Line&gt; Lines { get; init; } = [];
/// }
/// </code>
/// </example>
[AttributeUsage(AttributeTargets.Class, Inherited = false)]
public sealed class OrdTemplateAttribute : Attribute
{
}

/// <summary>Gives a property of an <see cref="OrdTemplateAttribute"/> class another name in the template, such as one with spaces or Swedish letters.</summary>
/// <param name="name">The name the template uses.</param>
[AttributeUsage(AttributeTargets.Property, Inherited = false)]
public sealed class TemplateNameAttribute(string name) : Attribute
{
    /// <summary>The name the template uses.</summary>
    public string Name { get; } = name;
}

/// <summary>Values for a template held by name: text, numbers and dates, true or false for a condition, lists of values for a section that repeats, and a <see cref="Replacement"/> for a picture or a table.</summary>
/// <example>
/// <code>
/// TemplateValues values = new TemplateValues()
///     .Set("Name", "Anna")
///     .Set("Total", 1200.5m)
///     .Set("Lines", [new TemplateValues().Set("Item", "Tea"), new TemplateValues().Set("Item", "Coffee")]);
/// </code>
/// </example>
public sealed class TemplateValues : ITemplateData
{
    private readonly Dictionary<string, object?> values = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Sets a value. Names are compared without regard to case.</summary>
    /// <returns>The same values, to go on setting.</returns>
    public TemplateValues Set(string name, object? value)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        values[name] = value;

        return this;
    }

    /// <summary>Values from a dictionary.</summary>
    public static TemplateValues From(IEnumerable<KeyValuePair<string, object?>> values)
    {
        ArgumentNullException.ThrowIfNull(values);
        TemplateValues made = new();

        foreach ((string name, object? value) in values)
        {
            made.Set(name, value);
        }

        return made;
    }

    /// <inheritdoc/>
    public bool TryGetValue(string name, out object? value)
    {
        return values.TryGetValue(name, out value);
    }
}

/// <summary>How a template is filled.</summary>
public sealed record TemplateOptions
{
    /// <summary>The options used when none are given.</summary>
    public static TemplateOptions Default { get; } = new();

    /// <summary>
    /// The culture numbers and dates are written in, as {{Total:N2}} or {{Date:d}} show them.
    /// The machine's culture when null, as in Word's own mail merge.
    /// </summary>
    public CultureInfo? Culture { get; init; }

    /// <summary>Fill Word's MERGEFIELD fields by their names. True by default.</summary>
    public bool MergeFields { get; init; } = true;

    /// <summary>Fill content controls by their tag, or their title when they have no tag. True by default.</summary>
    public bool ContentControls { get; init; } = true;
}

/// <summary>A placeholder, merge field or content control that had no value, and where it is.</summary>
/// <param name="Name">The name it asked for.</param>
/// <param name="Location">Where it is: the part, and the paragraph's text.</param>
public sealed record MissingValue(string Name, string Location);

/// <summary>What filling a template did.</summary>
/// <param name="Filled">How many placeholders, merge fields and content controls got a value.</param>
/// <param name="Missing">Those that had none, and became empty. Not an error, but a misspelt name shows here.</param>
public sealed record TemplateResult(int Filled, IReadOnlyList<MissingValue> Missing);

/// <summary>Looks names up through the values of the sections the template is in, innermost first.</summary>
internal sealed class TemplateScope
{
    private readonly TemplateScope? parent;
    private readonly ITemplateData? data;
    private readonly object? self;

    public TemplateScope(ITemplateData data)
    {
        this.data = data;
        self = data;
    }

    private TemplateScope(TemplateScope parent, object? item)
    {
        this.parent = parent;
        data = item as ITemplateData;
        self = item;
    }

    /// <summary>A scope for one item of a section: its names are looked up in it first.</summary>
    public TemplateScope For(object? item)
    {
        return new TemplateScope(this, item);
    }

    /// <summary>The value of a name: "." for the item itself, "Customer.Name" through nested values.</summary>
    public bool TryResolve(string name, out object? value)
    {
        value = null;

        if (name == ".")
        {
            value = self;

            return true;
        }

        string[] path = name.Split('.', StringSplitOptions.TrimEntries);

        if (TryFirst(path[0], out object? current) == false)
        {
            return false;
        }

        for (int index = 1; index < path.Length; index++)
        {
            if (current is not ITemplateData nested || nested.TryGetValue(path[index], out current) == false)
            {
                return false;
            }
        }

        value = current;

        return true;
    }

    /// <summary>The items a section repeats for: a list's items, once for a value that is true or not empty, never for false, null, "" or an empty list.</summary>
    public static IReadOnlyList<object?> Items(object? value)
    {
        List<object?> items = [];

        switch (value)
        {
            case null:
            case false:
            case "":
                break;

            case string:
                items.Add(value);

                break;

            case IEnumerable list:
                foreach (object? item in list)
                {
                    items.Add(item);
                }

                break;

            default:
                items.Add(value);

                break;
        }

        return items;
    }

    /// <summary>A value as text: formatted when it is IFormattable, with the format after the colon.</summary>
    public static string Format(object? value, string? format, CultureInfo culture)
    {
        return value switch
        {
            null => string.Empty,
            string text => text,
            bool flag => flag.ToString(culture),
            IFormattable formattable => formattable.ToString(format, culture),
            _ => value.ToString() ?? string.Empty,
        };
    }

    private bool TryFirst(string name, out object? value)
    {
        for (TemplateScope? scope = this; scope is not null; scope = scope.parent)
        {
            if (scope.data is not null && scope.data.TryGetValue(name, out value) == true)
            {
                return true;
            }
        }

        value = null;

        return false;
    }
}
