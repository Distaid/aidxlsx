#nullable disable

namespace Aid.Xlsx.Exceptions
{
    public class EmptyTableException : XlsxException
    {
        public EmptyTableException()
            : base()
        { }

        public EmptyTableException(String message)
            : base(message)
        { }

        public EmptyTableException(String message, Exception innerException)
            : base(message, innerException)
        { }
    }
}
