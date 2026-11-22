namespace Aid.Xlsx.Attributes;

/// <summary>Generates a reflection-free table row mapper for the annotated type.</summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct, Inherited = false)]
public sealed class XlsxTableAttribute(string? name = null) : Attribute
{
    /// <summary>The default Excel table name, if one is supplied.</summary>
    public string? Name { get; } = name;
}
