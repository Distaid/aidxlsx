namespace Aid.Xlsx.Model.RichText;

public interface IXlsxPhonetic : IEquatable<IXlsxPhonetic>
{
    string Text { get; }
    int Start { get; }
    int End { get; }
}
