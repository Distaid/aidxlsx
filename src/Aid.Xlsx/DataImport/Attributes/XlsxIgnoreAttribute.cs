namespace Aid.Xlsx.Attributes;

/// <summary>Excludes a field or property from a generated table row mapper.</summary>
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, Inherited = false)]
public sealed class XlsxIgnoreAttribute : Attribute;
