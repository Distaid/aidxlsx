#nullable disable

namespace Aid.Xlsx.Model.Exceptions;

public class EmptyTableException : XlsxException
{
    public EmptyTableException()
        : base()
    { }

    public EmptyTableException(string message)
        : base(message)
    { }

    public EmptyTableException(string message, Exception innerException)
        : base(message, innerException)
    { }
}
