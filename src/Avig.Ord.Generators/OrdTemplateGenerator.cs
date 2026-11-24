using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;

namespace Avig.Ord.Generators;

/// <summary>
/// Makes every [OrdTemplate] class an ITemplateData: a TryGetValue that hands out its public
/// properties by name (the property's name, or the one [TemplateName] gives), first as written
/// and then without regard to case. No reflection, so templates work with trimming and AOT.
/// </summary>
[Generator(LanguageNames.CSharp)]
public sealed class OrdTemplateGenerator : IIncrementalGenerator
{
    private const string AttributeName = "Avig.Ord.OrdTemplateAttribute";
    private const string NameAttributeName = "Avig.Ord.TemplateNameAttribute";

    private static readonly DiagnosticDescriptor NotPartial = new(
        "ORD001",
        "Template type must be partial",
        "'{0}' must be declared partial so that Avig.Ord can add its template values",
        "Avig.Ord",
        DiagnosticSeverity.Error,
        true);

    private static readonly DiagnosticDescriptor NotTopLevel = new(
        "ORD002",
        "Template type must be a top-level, non-generic type",
        "'{0}' is nested or generic. Declare template types at namespace level, without type parameters.",
        "Avig.Ord",
        DiagnosticSeverity.Error,
        true);

    /// <inheritdoc/>
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        IncrementalValuesProvider<Generated> types = context.SyntaxProvider.ForAttributeWithMetadataName(
            AttributeName,
            static (node, _) => node is ClassDeclarationSyntax or RecordDeclarationSyntax,
            static (attributeContext, _) => Generate((INamedTypeSymbol)attributeContext.TargetSymbol, (TypeDeclarationSyntax)attributeContext.TargetNode));

        context.RegisterSourceOutput(types, static (output, generated) =>
        {
            foreach (Diagnostic diagnostic in generated.Diagnostics)
            {
                output.ReportDiagnostic(diagnostic);
            }

            if (generated.Source is not null)
            {
                output.AddSource(generated.HintName, SourceText.From(generated.Source, Encoding.UTF8));
            }
        });
    }

    private static Generated Generate(INamedTypeSymbol type, TypeDeclarationSyntax syntax)
    {
        List<Diagnostic> diagnostics = [];
        Location location = syntax.Identifier.GetLocation();
        string hintName = type.ToDisplayString() + ".OrdTemplate.g.cs";

        if (syntax.Modifiers.Any(SyntaxKind.PartialKeyword) == false)
        {
            diagnostics.Add(Diagnostic.Create(NotPartial, location, type.Name));
        }

        if (type.ContainingType is not null || type.TypeParameters.Length > 0)
        {
            diagnostics.Add(Diagnostic.Create(NotTopLevel, location, type.Name));
        }

        if (diagnostics.Count > 0)
        {
            return new Generated(hintName, null, diagnostics);
        }

        return new Generated(hintName, Render(type, Properties(type)), diagnostics);
    }

    /// <summary>The public instance properties with a public getter, the base class's first, each name once.</summary>
    private static List<(string Name, string Property)> Properties(INamedTypeSymbol type)
    {
        List<(string Name, string Property)> properties = [];
        HashSet<string> seen = [];
        List<INamedTypeSymbol> chain = [];

        for (INamedTypeSymbol? current = type; current is not null && current.SpecialType != SpecialType.System_Object; current = current.BaseType)
        {
            chain.Insert(0, current);
        }

        foreach (INamedTypeSymbol current in chain)
        {
            foreach (IPropertySymbol property in current.GetMembers().OfType<IPropertySymbol>())
            {
                if (property.IsStatic == true || property.IsIndexer == true || property.DeclaredAccessibility != Accessibility.Public)
                {
                    continue;
                }

                if (property.GetMethod is null || property.GetMethod.DeclaredAccessibility != Accessibility.Public || property.Name == "EqualityContract")
                {
                    continue;
                }

                string name = property.Name;
                AttributeData? renamed = property.GetAttributes().FirstOrDefault(attribute => attribute.AttributeClass?.ToDisplayString() == NameAttributeName);

                if (renamed is not null && renamed.ConstructorArguments.Length == 1 && renamed.ConstructorArguments[0].Value is string given && given.Length > 0)
                {
                    name = given;
                }

                if (seen.Add(name) == true)
                {
                    properties.Add((name, property.Name));
                }
            }
        }

        return properties;
    }

    private static string Render(INamedTypeSymbol type, List<(string Name, string Property)> properties)
    {
        StringBuilder source = new();
        source.AppendLine("// <auto-generated/>");
        source.AppendLine("#nullable enable");
        source.AppendLine();

        if (type.ContainingNamespace.IsGlobalNamespace == false)
        {
            source.Append("namespace ").Append(type.ContainingNamespace.ToDisplayString()).AppendLine(";");
            source.AppendLine();
        }

        string kind = "partial class ";

        if (type.IsRecord == true)
        {
            kind = "partial record ";
        }

        source.Append(kind).Append(type.Name).AppendLine(" : global::Avig.Ord.ITemplateData");
        source.AppendLine("{");
        source.AppendLine("    bool global::Avig.Ord.ITemplateData.TryGetValue(string name, out object? value)");
        source.AppendLine("    {");
        source.AppendLine("        switch (name)");
        source.AppendLine("        {");

        foreach ((string name, string property) in properties)
        {
            source.Append("            case ").Append(SymbolDisplay.FormatLiteral(name, true)).AppendLine(":");
            source.Append("                value = this.").Append(property).AppendLine(";");
            source.AppendLine("                return true;");
        }

        source.AppendLine("        }");
        source.AppendLine();

        foreach ((string name, string property) in properties)
        {
            source.Append("        if (string.Equals(name, ").Append(SymbolDisplay.FormatLiteral(name, true)).AppendLine(", global::System.StringComparison.OrdinalIgnoreCase))");
            source.AppendLine("        {");
            source.Append("            value = this.").Append(property).AppendLine(";");
            source.AppendLine("            return true;");
            source.AppendLine("        }");
            source.AppendLine();
        }

        source.AppendLine("        value = null;");
        source.AppendLine("        return false;");
        source.AppendLine("    }");
        source.AppendLine("}");

        return source.ToString();
    }

    private sealed class Generated
    {
        public Generated(string hintName, string? source, List<Diagnostic> diagnostics)
        {
            HintName = hintName;
            Source = source;
            Diagnostics = diagnostics;
        }

        public string HintName { get; }

        public string? Source { get; }

        public List<Diagnostic> Diagnostics { get; }
    }
}
