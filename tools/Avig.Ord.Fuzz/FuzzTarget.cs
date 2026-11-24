namespace Avig.Ord.Fuzz;

/// <summary>
/// One thing to fuzz. <see cref="RunAsync"/> takes the raw input, and <see cref="Generate"/>
/// builds an input for smoke mode, where there is no coverage feedback to steer the mutations.
/// </summary>
/// <param name="Name">The name on the command line, such as <c>docx-part</c>.</param>
/// <param name="Description">One line for the usage text.</param>
/// <param name="RunAsync">The harness. Throws <see cref="InvariantViolationException"/> when the code answers wrong.</param>
/// <param name="Generate">Builds an input from a seeded random source.</param>
internal sealed record FuzzTarget(string Name, string Description, Func<byte[], Task> RunAsync, Func<Random, byte[]> Generate);

/// <summary>
/// Thrown by a harness when the code under test gave a wrong answer rather than an exception.
/// It is its own type so that it can never be mistaken for one of the exceptions the package
/// throws on purpose.
/// </summary>
public sealed class InvariantViolationException : Exception
{
    /// <summary>Creates the exception.</summary>
    public InvariantViolationException()
    {
    }

    /// <summary>Creates the exception.</summary>
    /// <param name="message">What did not hold.</param>
    public InvariantViolationException(string message)
        : base(message)
    {
    }

    /// <summary>Creates the exception.</summary>
    /// <param name="message">What did not hold.</param>
    /// <param name="innerException">The exception that showed it.</param>
    public InvariantViolationException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
