using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace Aid.Xlsx.Generators;

[Generator(LanguageNames.CSharp)]
public sealed class TableMapperGenerator : IIncrementalGenerator
{
    private const string Attributes = "Aid.Xlsx.Attributes.";
    private const string ColumnAttribute = Attributes + "XLColumnAttribute";
    private const string IgnoreAttribute = Attributes + "XlsxIgnoreAttribute";

    private static readonly DiagnosticDescriptor DuplicateColumn = Error("XLSX001", "Duplicate table column", "Table row '{0}' has duplicate column {1} '{2}'");
    private static readonly DiagnosticDescriptor InvalidMember = Error("XLSX002", "Invalid table column", "Member '{0}' must be a readable, public instance field or property with a supported cell value type");
    private static readonly DiagnosticDescriptor InvalidType = Error("XLSX003", "Invalid table row type", "Type '{0}' must be a non-generic, non-abstract class or struct accessible from the generated mapper");
    private static readonly DiagnosticDescriptor NoColumns = Error("XLSX005", "Empty table row", "Table row '{0}' has no readable columns");

    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        var rows = context.SyntaxProvider.ForAttributeWithMetadataName(Attributes + "XlsxTableAttribute",
            static (node, _) => node is TypeDeclarationSyntax,
            static (syntax, _) => (Type: (INamedTypeSymbol)syntax.TargetSymbol, Attribute: syntax.Attributes[0]));

        context.RegisterSourceOutput(rows, static (context, row) => Emit(context, row.Type, row.Attribute));
    }

    private static void Emit(SourceProductionContext context, INamedTypeSymbol type, AttributeData attribute)
    {
        var typeName = type.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
        if (type.IsStatic || type.IsAbstract || type.IsRefLikeType || !IsAccessible(type))
        {
            context.ReportDiagnostic(Diagnostic.Create(InvalidType, type.Locations.FirstOrDefault(), typeName));
            return;
        }

        var members = new List<Column>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var invalid = false;
        for (var current = type; current is not null && current.SpecialType != SpecialType.System_Object; current = current.BaseType)
        {
            foreach (var member in current.GetMembers())
            {
                if (member is not (IFieldSymbol or IPropertySymbol) || member is IFieldSymbol { IsImplicitlyDeclared: true })
                    continue;

                var attributes = member.GetAttributes();
                var column = attributes.FirstOrDefault(x => x.AttributeClass?.ToDisplayString() == ColumnAttribute);
                var settings = column?.NamedArguments.ToDictionary(x => x.Key, x => x.Value.Value, StringComparer.Ordinal);
                var isPublic = member.DeclaredAccessibility == Accessibility.Public;
                if ((!isPublic && column is null) || !seen.Add(member.Name))
                    continue;
                if (attributes.Any(x => x.AttributeClass?.ToDisplayString() == IgnoreAttribute)
                    || GetSetting(settings, "Ignore") is true)
                    continue;

                var memberType = member switch
                {
                    IFieldSymbol field when isPublic && !field.IsStatic && !field.IsFixedSizeBuffer => field.Type,
                    IPropertySymbol { GetMethod.DeclaredAccessibility: Accessibility.Public, IsIndexer: false } property
                        when isPublic && !property.IsStatic => property.Type,
                    _ => null
                };
                if (memberType is null)
                {
                    if (column is not null)
                    {
                        context.ReportDiagnostic(Diagnostic.Create(InvalidMember, member.Locations.FirstOrDefault(), member.Name));
                        invalid = true;
                    }
                    continue;
                }

                var accessor = GetAccessor(memberType, current, member.Name);
                if (accessor is null)
                {
                    context.ReportDiagnostic(Diagnostic.Create(InvalidMember, member.Locations.FirstOrDefault(), member.Name));
                    invalid = true;
                    continue;
                }

                var header = GetSetting(settings, "Header") as string;
                var order = (int?)GetSetting(settings, "Order");
                members.Add(new Column(member, string.IsNullOrWhiteSpace(header) ? member.Name : header!, order,
                    GetSetting(settings, "Format") as string, accessor));
            }
        }

        if (invalid)
            return;
        if (members.Count == 0)
        {
            context.ReportDiagnostic(Diagnostic.Create(NoColumns, type.Locations.FirstOrDefault(), typeName));
            return;
        }

        var headers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var orders = new HashSet<int>();
        foreach (var member in members)
        {
            var location = member.Symbol.Locations.FirstOrDefault();
            if (!headers.Add(member.Header))
            {
                context.ReportDiagnostic(Diagnostic.Create(DuplicateColumn, location, typeName, "header", member.Header));
                invalid = true;
            }
            if (member.Order is int order && !orders.Add(order))
            {
                context.ReportDiagnostic(Diagnostic.Create(DuplicateColumn, location, typeName, "order", order));
                invalid = true;
            }
        }
        if (invalid)
            return;

        var sorted = members.OrderBy(x => x.Order ?? int.MaxValue).ToArray();
        var tableName = attribute.ConstructorArguments.FirstOrDefault().Value as string;
        using var hash = SHA256.Create();
        var identifier = "TableMapper_" + BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(typeName))).Replace("-", "");
        var source = new StringBuilder();
        source.AppendLine("// <auto-generated/>")
            .AppendLine("#nullable enable")
            .AppendLine("namespace Aid.Xlsx.Generated")
            .AppendLine("{")
            .Append("    internal static class ").AppendLine(identifier)
            .AppendLine("    {")
            .AppendLine("        [global::System.Runtime.CompilerServices.ModuleInitializer]")
            .AppendLine("        internal static void Register()")
            .AppendLine("        {")
            .Append("            global::Aid.Xlsx.InsertData.TableMapperRegistry.Register<").Append(typeName).AppendLine(">(")
            .Append("                ").Append(Literal(tableName)).AppendLine(",")
            .Append("                new string[] { ").Append(string.Join(", ", sorted.Select(x => Literal(x.Header)))).AppendLine(" },")
            .Append("                new string?[] { ").Append(string.Join(", ", sorted.Select(x => Literal(x.Format)))).AppendLine(" },")
            .AppendLine("                static (item, column) => column switch")
            .AppendLine("                {");
        for (var i = 0; i < sorted.Length; i++)
            source.Append("                    ").Append(i.ToString(CultureInfo.InvariantCulture)).Append(" => ").Append(sorted[i].Accessor).AppendLine(",");

        source.AppendLine("                    _ => throw new global::System.ArgumentOutOfRangeException(nameof(column))")
            .AppendLine("                });")
            .AppendLine("        }")
            .AppendLine("    }")
            .AppendLine("}");
        context.AddSource(identifier + ".g.cs", source.ToString());
    }

    private static bool IsAccessible(INamedTypeSymbol type)
    {
        for (var current = type; current is not null; current = current.ContainingType)
        {
            if (current.Arity != 0 || current.IsFileLocal || current.DeclaredAccessibility is not
                (Accessibility.Public or Accessibility.Internal or Accessibility.ProtectedOrInternal))
                return false;
        }
        return true;
    }

    private static string? GetAccessor(ITypeSymbol type, INamedTypeSymbol declaringType, string member)
    {
        var access = $"(({declaringType.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat)})item).@{member}";
        var nullable = type is INamedTypeSymbol { OriginalDefinition.SpecialType: SpecialType.System_Nullable_T };
        if (nullable)
            type = ((INamedTypeSymbol)type).TypeArguments[0];

        var value = nullable ? "value" : access;
        var expression = type.SpecialType switch
        {
            SpecialType.System_Boolean or SpecialType.System_String or SpecialType.System_DateTime or
            SpecialType.System_SByte or SpecialType.System_Byte or SpecialType.System_Int16 or SpecialType.System_UInt16 or
            SpecialType.System_Int32 or SpecialType.System_UInt32 or SpecialType.System_Int64 or SpecialType.System_UInt64 or
            SpecialType.System_Single or SpecialType.System_Double or SpecialType.System_Decimal => value,
            SpecialType.System_Char => value + ".ToString()",
            _ when type.TypeKind == TypeKind.Enum && type.ToDisplayString() != "Aid.Xlsx.XLError" => value + ".ToString()",
            _ => type.ToDisplayString() switch
            {
                "Aid.Xlsx.XLCellValue" or "Aid.Xlsx.XLError" or "Aid.Xlsx.Blank" or "System.TimeSpan" => value,
                "System.DateOnly" => value + ".ToDateTime(global::System.TimeOnly.MinValue)",
                "System.TimeOnly" => value + ".ToTimeSpan()",
                "System.DateTimeOffset" => value + ".DateTime",
                "System.Guid" => value + ".ToString()",
                _ => null
            }
        };

        if (expression is null || !nullable)
            return expression;

        return $"{access} is {{ }} value ? (global::Aid.Xlsx.XLCellValue)({expression}) : default";
    }

    private static object? GetSetting(Dictionary<string, object?>? settings, string key)
        => settings is not null && settings.TryGetValue(key, out var value) ? value : null;

    private static string Literal(string? value) => value is null ? "null" : SymbolDisplay.FormatLiteral(value, quote: true);

    private static DiagnosticDescriptor Error(string id, string title, string message)
        => new(id, title, message, "Aid.Xlsx", DiagnosticSeverity.Error, isEnabledByDefault: true);

    private sealed class Column(ISymbol symbol, string header, int? order, string? format, string accessor)
    {
        public ISymbol Symbol { get; } = symbol;
        public string Header { get; } = header;
        public int? Order { get; } = order;
        public string? Format { get; } = format;
        public string Accessor { get; } = accessor;
    }
}
