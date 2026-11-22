namespace Aid.Xlsx.RichText
{
    public interface IXlsxPhonetic : IEquatable<IXlsxPhonetic>
    {
        String Text { get; }
        Int32 Start { get; }
        Int32 End { get; }
    }
}
