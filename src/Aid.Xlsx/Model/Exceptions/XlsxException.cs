#nullable disable

namespace Aid.Xlsx.Model.Exceptions;

public abstract class XlsxException : Exception
{
    protected XlsxException()
        : base()
    { }

    protected XlsxException(String message)
        : base(message)
    { }

    protected XlsxException(String message, Exception innerException)
        : base(message, innerException)
    { }
}
