namespace Aid.Xlsx.Styles.Formatting;

/// <summary>
/// A key to a named style that is used to connect immutable <see cref="XlsxCellFormatValue"/> to
/// a changeable <see cref="XlsxCellStyleValue"/>. API methods that needs access to a resolved format
/// value should pass <see cref="XlsxWorkbookStyles.DefaultFormat"/> as a parameter.
/// </summary>
internal readonly record struct StyleId(int Value) : IEquatable<int>, IComparable<StyleId>
{
    public static implicit operator StyleId(int v) => new(v);

    bool IEquatable<int>.Equals(int other)
    {
        return Value == other;
    }

    public int CompareTo(StyleId other)
    {
        return Value.CompareTo(other.Value);
    }
}
