#nullable disable

namespace Aid.Xlsx.Model.Exceptions;

public abstract class XlsxException : Exception
{
    protected XlsxException()
        : base()
    { }

    protected XlsxException(string message)
        : base(message)
    { }

    protected XlsxException(string message, Exception innerException)
        : base(message, innerException)
    { }
}
