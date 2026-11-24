// Makes README.md from README.template.md: every @@name@@ becomes the code of
// "#region example: name" in tests/Avig.Ord.Tests/ReadmeExamples.cs, dedented, so that the
// README shows exactly the code the tests run. Run from the repo root: dotnet run tools/readme.cs
using System.Text.RegularExpressions;

string root = Directory.GetCurrentDirectory();
string source = File.ReadAllText(Path.Combine(root, "tests", "Avig.Ord.Tests", "ReadmeExamples.cs")).Replace("\r\n", "\n", StringComparison.Ordinal);
string template = File.ReadAllText(Path.Combine(root, "README.template.md")).Replace("\r\n", "\n", StringComparison.Ordinal);
Dictionary<string, string> regions = [];

foreach (Match match in Regex.Matches(source, @"^[ \t]*#region example: (?<name>[\w-]+)\n(?<code>.*?)^[ \t]*#endregion", RegexOptions.Multiline | RegexOptions.Singleline))
{
    string[] lines = match.Groups["code"].Value.TrimEnd().Split('\n');
    int indent = lines.Where(line => line.Trim().Length > 0).Select(line => line.Length - line.TrimStart().Length).DefaultIfEmpty(0).Min();
    regions[match.Groups["name"].Value] = string.Join("\n", lines.Select(line => Dedent(line, indent)));
}

string readme = Regex.Replace(template, @"@@(?<name>[\w-]+)@@", match =>
{
    string name = match.Groups["name"].Value;

    if (regions.TryGetValue(name, out string? code) == false)
    {
        throw new InvalidOperationException($"No #region example: {name} in ReadmeExamples.cs.");
    }

    return code;
});

File.WriteAllText(Path.Combine(root, "README.md"), readme.Replace("\n", "\r\n", StringComparison.Ordinal));
Console.WriteLine($"README.md written with {regions.Count} examples.");

static string Dedent(string line, int indent)
{
    if (line.Length >= indent)
    {
        return line[indent..];
    }

    return line.TrimStart();
}
