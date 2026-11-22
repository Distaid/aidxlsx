using System.Globalization;
using System.Reflection;
using System.Text.Json;

internal sealed class ApiSurface(Assembly assembly)
{
    private const BindingFlags Declared = BindingFlags.DeclaredOnly | BindingFlags.Public | BindingFlags.NonPublic |
                                          BindingFlags.Instance | BindingFlags.Static;
    private readonly NullabilityInfoContext _nullability = new();

    public static HashSet<string> Read(Assembly assembly) => new ApiSurface(assembly).Read();

    private HashSet<string> Read()
    {
        var declarations = new HashSet<string>(StringComparer.Ordinal);
        foreach (var type in assembly.GetExportedTypes())
        {
            var owner = TypeName(type);
            var kind = type.IsInterface ? "interface" : type.IsEnum ? "enum" : type.IsValueType ? "struct" : "class";
            var attributes = Attributes(type.CustomAttributes);
            var modifiers = string.Join(' ', new[] { type.IsAbstract && !type.IsInterface ? "abstract" : "", type.IsSealed ? "sealed" : "" }.Where(s => s.Length > 0));
            var baseType = type.IsEnum ? TypeName(type.GetEnumUnderlyingType()) : type.BaseType is { } parent ? TypeName(parent) : "-";
            declarations.Add($"type {owner} {kind} {modifiers} : {baseType} {Constraints(type.GetGenericArguments())}{attributes}".Trim());
            foreach (var implemented in type.GetInterfaces())
                declarations.Add($"interface {owner} : {TypeName(implemented)}");

            var properties = type.GetProperties(Declared);
            var events = type.GetEvents(Declared);
            var accessors = properties.SelectMany(p => p.GetAccessors(true))
                .Concat(events.SelectMany(e => new[] { e.AddMethod, e.RemoveMethod, e.RaiseMethod }.OfType<MethodInfo>()))
                .ToHashSet();

            foreach (var member in type.GetMembers(Declared))
            {
                var signature = member switch
                {
                    ConstructorInfo ctor when Visible(ctor) => $"ctor {Method(ctor)}({Parameters(ctor.GetParameters())})",
                    MethodInfo method when Visible(method) && !accessors.Contains(method) =>
                        $"method {Method(method)} {TypeName(method.ReturnType, _nullability.Create(method.ReturnParameter))} {method.Name}{GenericArguments(method.GetGenericArguments())}({Parameters(method.GetParameters())}){Constraints(method.GetGenericArguments())}{Attributes(method.ReturnParameter.CustomAttributes)}",
                    PropertyInfo p when Visible(p.GetMethod) || Visible(p.SetMethod) =>
                        $"property {TypeName(p.PropertyType, _nullability.Create(p))} {p.Name}({Parameters(p.GetIndexParameters())}){Accessor(p.GetMethod, "get")}{Accessor(p.SetMethod, "set")}",
                    EventInfo e when Visible(e.AddMethod) || Visible(e.RemoveMethod) =>
                        $"event {TypeName(e.EventHandlerType!, _nullability.Create(e))} {e.Name}{Accessor(e.AddMethod, "add")}{Accessor(e.RemoveMethod, "remove")}",
                    FieldInfo field when field.IsPublic || field.IsFamily || field.IsFamilyOrAssembly =>
                        $"field {Access(field.IsPublic, field.IsFamily)} {(field.IsStatic ? "static " : "")}{(field.IsInitOnly ? "readonly " : "")}{TypeName(field.FieldType, _nullability.Create(field))} {field.Name}{(field.IsLiteral ? " = " + Literal(field.GetRawConstantValue()) : "")}",
                    _ => null,
                };
                if (signature is not null)
                    declarations.Add($"{owner} {signature}{Attributes(member.CustomAttributes)}");
            }
        }

        return declarations;
    }

    private static bool Visible(MethodBase? method) => method is { IsPublic: true } or { IsFamily: true } or { IsFamilyOrAssembly: true };

    private static string Access(bool isPublic, bool isProtected) => isPublic ? "public" : isProtected ? "protected" : "protected-internal";

    private static string Method(MethodBase method) => Access(method.IsPublic, method.IsFamily)
        + (method.IsStatic ? " static" : "")
        + (method.IsAbstract ? " abstract" : method.IsVirtual ? " virtual" : "")
        + (method.IsFinal ? " final" : "");

    private static string Accessor(MethodInfo? method, string name) => method is not null && Visible(method)
        ? $" {name}({Method(method)}{(method.ReturnParameter.GetRequiredCustomModifiers().Any(m => m.Name == "IsExternalInit") ? " init" : "")})"
        : "";

    private string Parameters(ParameterInfo[] parameters) => string.Join(", ", parameters.Select(p =>
        $"{(p.IsOut ? "out " : p.IsIn ? "in " : "")}{TypeName(p.ParameterType, _nullability.Create(p))} {p.Name}{(p.HasDefaultValue ? " = " + Literal(p.RawDefaultValue) : "")}{Attributes(p.CustomAttributes)}"));

    private string Constraints(Type[] types) => string.Concat(types.Where(type => type.IsGenericParameter).Select(type =>
        $" [{TypeName(type)} {(int)type.GenericParameterAttributes} {string.Join(",", type.GetGenericParameterConstraints().Select(constraint => TypeName(constraint)).Order(StringComparer.Ordinal))}]"));

    private string GenericArguments(Type[] types) => types.Length == 0 ? "" : $"<{string.Join(",", types.Select(type => TypeName(type)))}>";

    private string TypeName(Type type, NullabilityInfo? nullability = null)
    {
        if (type.IsByRef)
            return "ref " + TypeName(type.GetElementType()!, nullability);
        var suffix = nullability?.ReadState == NullabilityState.Nullable && !type.IsValueType ? "?" : "";
        if (type.IsArray)
            return TypeName(type.GetElementType()!, nullability?.ElementType) + "[" + new string(',', type.GetArrayRank() - 1) + "]" + suffix;
        if (type.IsGenericParameter)
            return (type.DeclaringMethod is null ? "!" : "!!") + type.GenericParameterPosition + suffix;

        var definition = type.IsConstructedGenericType ? type.GetGenericTypeDefinition() : type;
        var name = definition.FullName ?? definition.Name;

        var arguments = type.IsGenericType
            ? "<" + string.Join(",", type.GetGenericArguments().Select((argument, i) => TypeName(argument, nullability?.GenericTypeArguments.ElementAtOrDefault(i)))) + ">"
            : "";
        return name + arguments + suffix;
    }

    private static string Attributes(IEnumerable<CustomAttributeData> attributes) => string.Concat(attributes
        .Where(attr => attr.AttributeType.Name is "ObsoleteAttribute" or "EditorBrowsableAttribute" or "FlagsAttribute"
            or "AttributeUsageAttribute" or "DynamicAttribute" or "ParamArrayAttribute" or "ExtensionAttribute")
        .Select(attr => $" [{attr.AttributeType.Name}({string.Join(",", attr.ConstructorArguments.Select(AttributeValue))}{string.Concat(attr.NamedArguments.OrderBy(a => a.MemberName, StringComparer.Ordinal).Select(a => $",{a.MemberName}={AttributeValue(a.TypedValue)}"))})]")
        .Order(StringComparer.Ordinal));

    private static string AttributeValue(CustomAttributeTypedArgument argument) => argument.Value is IEnumerable<CustomAttributeTypedArgument> elements
        ? $"[{string.Join(",", elements.Select(AttributeValue))}]"
        : Literal(argument.Value);

    private static string Literal(object? value) => value switch
    {
        null => "null",
        string text => JsonSerializer.Serialize(text),
        char character => JsonSerializer.Serialize(character),
        _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? "null",
    };
}
