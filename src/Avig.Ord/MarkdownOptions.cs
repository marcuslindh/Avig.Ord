namespace Avig.Ord;

/// <summary>How a document is written as Markdown.</summary>
public sealed record MarkdownOptions
{
    /// <summary>The options used when none are given.</summary>
    public static MarkdownOptions Default { get; } = new();

    /// <summary>
    /// The folder picture links point into, relative to the Markdown file: "media" by default,
    /// which gives <c>![alt](media/image1.png)</c>. <see cref="OrdDocument.SaveMarkdownAsync"/>
    /// saves the pictures there.
    /// </summary>
    public string PictureFolder { get; init; } = "media";

    /// <summary>
    /// Makes the link for a picture instead of <see cref="PictureFolder"/> and the picture's
    /// file name, for pictures stored somewhere else, such as a web server. Null by default.
    /// </summary>
    public Func<Picture, string>? PictureLink { get; init; }

    /// <summary>Write the text of text boxes after the paragraph they are anchored in. True by default.</summary>
    public bool TextBoxes { get; init; } = true;
}
