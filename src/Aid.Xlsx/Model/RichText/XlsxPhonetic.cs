namespace Aid.Xlsx.Model.RichText;

internal class XlsxPhonetic(string text, int start, int end) : IXlsxPhonetic
{
    public string Text { get; } = text;
    public int Start { get; } = start;
    public int End { get; } = end;

    public bool Equals(IXlsxPhonetic? other)
    {
        if (other is null)
            return false;

        if (ReferenceEquals(this, other))
            return true;

        return Text == other.Text && Start == other.Start && End == other.End;
    }
}
