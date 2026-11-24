using System.Diagnostics;
using System.Globalization;
using SharpFuzz;

namespace Avig.Ord.Fuzz;

/// <summary>
/// Fuzz harness for Avig.Ord. Runs a target under libFuzzer (through <c>fuzz.ps1</c>), as
/// a seeded smoke run without coverage feedback, or on one saved input to show the exception.
/// </summary>
internal static class Program
{
    /// <summary>An input that takes longer than this is saved, because a hang is a finding too.</summary>
    private static readonly TimeSpan SlowInput = TimeSpan.FromSeconds(2);

    private static readonly FuzzTarget[] Targets =
    [
        PartTarget.Create(),
        ZipTarget.Create(),
        FlatTarget.Create(),
        RoundTripTarget.Create(),
        BodyTargets.Render(),
        BodyTargets.Events(),
        BodyTargets.ReaderModel(),
        BodyTargets.Revisions(),
        ValueTargets.ImageSize(),
        ValueTargets.Format(),
    ];

    private static async Task<int> Main(string[] args)
    {
        if (args.Length == 0)
        {
            PrintUsage();

            return 1;
        }

        string name = args[0].ToLowerInvariant();
        FuzzTarget[] selected;

        if (name == "all")
        {
            selected = Targets;
        }
        else
        {
            selected = Targets.Where(target => target.Name == name).ToArray();
        }

        if (selected.Length == 0)
        {
            Console.Error.WriteLine($"Unknown target '{args[0]}'.");
            PrintUsage();

            return 1;
        }

        if (TryGetOption(args, "--repro", out string? reproFile) == true)
        {
            if (TryGetOption(args, "--dump", out string? dumpDirectory) == true)
            {
                Dump.Directory = Path.GetFullPath(dumpDirectory);
            }

            return await ReproAsync(selected[0], reproFile);
        }

        int seed = Random.Shared.Next();

        if (TryGetOption(args, "--seed", out string? seedText) == true)
        {
            seed = int.Parse(seedText, CultureInfo.InvariantCulture);
        }

        if (TryGetOption(args, "--write-corpus", out string? countText) == true)
        {
            return WriteCorpus(selected, int.Parse(countText, CultureInfo.InvariantCulture), seed);
        }

        if (HasFlag(args, "--smoke") == true)
        {
            int iterations = 20_000;

            if (TryGetOption(args, "--iter", out string? iterationText) == true)
            {
                iterations = int.Parse(iterationText, CultureInfo.InvariantCulture);
            }

            return await SmokeAsync(selected, iterations, seed, HasFlag(args, "--trace"));
        }

        if (selected.Length != 1)
        {
            Console.Error.WriteLine("libFuzzer mode takes exactly one target.");

            return 1;
        }

        FuzzTarget only = selected[0];

        // libFuzzer calls a synchronous callback once per input and there is no caller to hand
        // a task to, so this is the one place the harness blocks.
#pragma warning disable AVIG0014
        Fuzzer.LibFuzzer.Run(span => only.RunAsync(span.ToArray()).GetAwaiter().GetResult());
#pragma warning restore AVIG0014

        return 0;
    }

    private static async Task<int> SmokeAsync(FuzzTarget[] targets, int iterations, int seed, bool trace)
    {
        Console.WriteLine($"Smoke fuzz: {iterations} inputs per target, seed {seed}");
        Console.WriteLine();
        int totalFindings = 0;

        for (int targetIndex = 0; targetIndex < targets.Length; targetIndex++)
        {
            FuzzTarget target = targets[targetIndex];
            Random random = new(seed + targetIndex);
            string findings = Path.Combine(ProjectDirectory(), "findings", target.Name);
            int passed = 0;
            int rejected = 0;
            int found = 0;
            long started = Stopwatch.GetTimestamp();

            for (int iteration = 0; iteration < iterations && found < 5; iteration++)
            {
                byte[] input = target.Generate(random);
                string inputName = string.Create(CultureInfo.InvariantCulture, $"{seed:x8}-{iteration:x6}.bin");

                if (trace == true)
                {
                    // A stack overflow ends the process without running any catch, so the only
                    // way to keep the input is to write it before it runs.
                    Directory.CreateDirectory(findings);
                    await File.WriteAllBytesAsync(Path.Combine(findings, "last-input.bin"), input);
                }

                long inputStarted = Stopwatch.GetTimestamp();

                int refusalsBefore = Harness.Refusals;

                try
                {
                    // Nothing is swallowed here: each target catches the refusals the package documents, and checks that they have a kind.
                    await target.RunAsync(input);

                    if (Harness.Refusals == refusalsBefore)
                    {
                        passed++;
                    }
                    else
                    {
                        rejected++;
                    }
                }
                catch (Exception exception)
                {
                    found++;
                    string path = await SaveAsync(findings, "crash-" + inputName, input);
                    Console.Error.WriteLine($"  [{target.Name}] {exception.GetType().Name}: {Shorten(exception.Message, 160)}");
                    Console.Error.WriteLine($"      input of {input.Length} bytes -> {path}");
                }

                TimeSpan inputTime = Stopwatch.GetElapsedTime(inputStarted);

                if (inputTime > SlowInput)
                {
                    found++;
                    string path = await SaveAsync(findings, "slow-" + inputName, input);
                    Console.Error.WriteLine($"  [{target.Name}] slow input, {inputTime.TotalSeconds:0.0} s -> {path}");
                }
            }

            TimeSpan elapsed = Stopwatch.GetElapsedTime(started);
            string status = "OK  ";

            if (found > 0)
            {
                status = "FAIL";
            }

            Console.WriteLine(string.Create(
                CultureInfo.InvariantCulture,
                $"  {status} {target.Name,-14} passed={passed,-7} rejected={rejected,-7} findings={found,-3} {elapsed.TotalSeconds,6:0.0}s"));
            totalFindings += found;
        }

        Console.WriteLine();

        if (Harness.KnownFlatDepthIssues > 0)
        {
            Console.WriteLine($"Known: {Harness.KnownFlatDepthIssues} Flat OPC saves hit the depth limit of 100 in Avig.OpenXml's writer and were saved as zip (README, Findings).");
        }

        if (totalFindings > 0)
        {
            Console.WriteLine($"FAIL: {totalFindings} findings, saved under findings/.");

            return 1;
        }

        Console.WriteLine("OK: no findings.");

        return 0;
    }

    /// <summary>Runs one saved input without swallowing anything, and prints the whole exception.</summary>
    private static async Task<int> ReproAsync(FuzzTarget target, string file)
    {
        byte[] input = await File.ReadAllBytesAsync(file);
        Console.WriteLine($"Repro: {target.Name}, {file}, {input.Length} bytes");

        try
        {
            await target.RunAsync(input);
            Console.WriteLine("(no exception)");

            return 0;
        }
        catch (Exception exception)
        {
            Console.WriteLine();
            Console.WriteLine(exception.ToString());

            return 1;
        }
    }

    /// <summary>Writes generated inputs to corpus/&lt;target&gt;/ as starting points for libFuzzer.</summary>
    private static int WriteCorpus(FuzzTarget[] targets, int count, int seed)
    {
        foreach (FuzzTarget target in targets)
        {
            Random random = new(seed);
            string directory = Path.Combine(ProjectDirectory(), "corpus", target.Name);
            Directory.CreateDirectory(directory);

            for (int index = 0; index < count; index++)
            {
                string name = string.Create(CultureInfo.InvariantCulture, $"seed-{index + 1:000}.bin");
                File.WriteAllBytes(Path.Combine(directory, name), target.Generate(random));
            }

            Console.WriteLine($"{target.Name}: {count} inputs in {directory}");
        }

        return 0;
    }

    private static async Task<string> SaveAsync(string directory, string name, byte[] input)
    {
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, name);
        await File.WriteAllBytesAsync(path, input);

        return path;
    }

    /// <summary>The folder with the project file, so that findings and corpus land next to the code however the harness was started.</summary>
    private static string ProjectDirectory()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);

        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Avig.Ord.Fuzz.csproj")) == true)
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        return Environment.CurrentDirectory;
    }

    private static bool HasFlag(string[] args, string flag)
    {
        return args.Any(argument => argument.Equals(flag, StringComparison.OrdinalIgnoreCase));
    }

    private static bool TryGetOption(string[] args, string option, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out string? value)
    {
        int index = Array.FindIndex(args, argument => argument.Equals(option, StringComparison.OrdinalIgnoreCase));
        value = null;

        if (index < 0 || index + 1 >= args.Length)
        {
            return false;
        }

        value = args[index + 1];

        return true;
    }

    private static string Shorten(string text, int length)
    {
        string line = text.ReplaceLineEndings(" ");

        if (line.Length <= length)
        {
            return line;
        }

        return line[..length] + "...";
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage: Avig.Ord.Fuzz <target|all> [--smoke [--iter N] [--seed S] [--trace]] [--repro FILE [--dump DIR]] [--write-corpus N]");
        Console.WriteLine();
        Console.WriteLine("Without --smoke the target runs under libFuzzer; start it with fuzz.ps1.");
        Console.WriteLine();

        foreach (FuzzTarget target in Targets)
        {
            Console.WriteLine($"  {target.Name,-14} {target.Description}");
        }
    }
}
