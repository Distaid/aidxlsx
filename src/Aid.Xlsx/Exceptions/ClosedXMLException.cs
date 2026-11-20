#nullable disable

using System;

namespace Aid.Xlsx.Exceptions
{
    public abstract class ClosedXMLException : Exception
    {
        protected ClosedXMLException()
            : base()
        { }

        protected ClosedXMLException(String message)
            : base(message)
        { }

        protected ClosedXMLException(String message, Exception innerException)
            : base(message, innerException)
        { }
    }
}
