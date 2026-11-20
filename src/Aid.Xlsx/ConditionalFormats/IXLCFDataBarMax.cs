#nullable disable

using System;

namespace Aid.Xlsx
{
    public interface IXLCFDataBarMax
    {
        void Maximum(XLCFContentType type, String value);
        void Maximum(XLCFContentType type, Double value);
        void HighestValue();
    }
}
