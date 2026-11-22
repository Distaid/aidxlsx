using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Styles;

/// <summary>
/// A font name, two font names are equal when they are case insensitive equal. It is a custom
/// class because that way <see cref="XlsxFontFormatValue"/> and other structures don't have to implement
/// custom hash code and equality methods.
/// </summary>
internal readonly record struct XlsxFontName : IEquatable<string>
{
    private const StringComparison Comparison = StringComparison.OrdinalIgnoreCase;

    private XlsxFontName(string text)
    {
        // Spec says at most 31 chars, Excel also tries to repair workbook when value is longer.
        if (string.IsNullOrWhiteSpace(text) || text.Length > 31)
            throw new ArgumentException("Font name can't be empty and must be less than 32 characters long.", nameof(text));

        Text = text;
    }

    public string Text { get; }

    public bool Equals(string? other)
    {
        return string.Equals(Text, other, Comparison);
    }

    public override int GetHashCode()
    {
        return Text.GetHashCode(Comparison);
    }

    public bool Equals(XlsxFontName other)
    {
        return Equals(other.Text);
    }

    public static implicit operator XlsxFontName(string text) => new(text);
}
