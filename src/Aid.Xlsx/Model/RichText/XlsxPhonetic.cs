namespace Aid.Xlsx.Model.RichText;

internal class XlsxPhonetic(String text, Int32 start, Int32 end) : IXlsxPhonetic
{
    public String Text { get; } = text;
    public Int32 Start { get; } = start;
    public Int32 End { get; } = end;

    public bool Equals(IXlsxPhonetic? other)
    {
        if (other is null)
            return false;

        if (ReferenceEquals(this, other))
            return true;

        return Text == other.Text && Start == other.Start && End == other.End;
    }
}
