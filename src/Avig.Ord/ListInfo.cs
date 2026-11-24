namespace Avig.Ord;

/// <summary>A paragraph's place in a list: its level and the label Word shows before it.</summary>
/// <param name="Level">The level, 0 for the outermost.</param>
/// <param name="Label">The label as Word shows it: "1.", "2.3", "a)", "iv." or a bullet such as •.</param>
/// <param name="IsBullet">True for a bulleted list, false for a numbered one.</param>
/// <param name="NumberingId">The w:numId of the list, which tells lists apart.</param>
public sealed record ListInfo(int Level, string Label, bool IsBullet, string NumberingId)
{
    /// <summary>The item's number at its level, counted as Word counts: 1, 2, 3 … also for letters and Roman numerals.</summary>
    public int Number { get; init; }

    /// <summary>Word's name for the number format: decimal, lowerLetter, upperRoman, bullet …</summary>
    public string NumberFormat { get; init; } = "decimal";
}

/// <summary>How tracked changes are read.</summary>
public enum RevisionView
{
    /// <summary>As if every change had been accepted: inserted text is there, deleted text is not. The default.</summary>
    Accepted,

    /// <summary>As if every change had been rejected: the text before the changes.</summary>
    Original,

    /// <summary>Both inserted and deleted text, each run marked with its <see cref="Run.Revision"/>.</summary>
    Marked,
}

/// <summary>Whether a run is part of a tracked change.</summary>
public enum RevisionKind
{
    /// <summary>Not part of a tracked change.</summary>
    None,

    /// <summary>Inserted while changes were tracked.</summary>
    Inserted,

    /// <summary>Deleted while changes were tracked.</summary>
    Deleted,
}

/// <summary>How a document is read.</summary>
public sealed record ReadOptions
{
    /// <summary>The options used when none are given.</summary>
    public static ReadOptions Default { get; } = new();

    /// <summary>How tracked changes are read. <see cref="RevisionView.Accepted"/> by default.</summary>
    public RevisionView Revisions { get; init; } = RevisionView.Accepted;

    /// <summary>The limits on the file. <see cref="OrdLimits.Default"/> by default.</summary>
    public OrdLimits Limits { get; init; } = OrdLimits.Default;
}
