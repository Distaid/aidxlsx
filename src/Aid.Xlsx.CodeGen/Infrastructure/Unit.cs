namespace Aid.Xlsx.CodeGen.Infrastructure;

/// <summary>
/// A type to use in generics instead of <c>void</c>.
/// </summary>
internal readonly struct Unit
{
    public static readonly Unit Value = new();
};
