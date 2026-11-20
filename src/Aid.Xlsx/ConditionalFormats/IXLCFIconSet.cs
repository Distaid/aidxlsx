#nullable disable

using System;

namespace Aid.Xlsx
{
    public enum XLCFIconSetOperator { GreaterThan, EqualOrGreaterThan }
    public interface IXLCFIconSet
    {
        IXLCFIconSet AddValue(XLCFIconSetOperator setOperator, String value, XLCFContentType type);
        IXLCFIconSet AddValue(XLCFIconSetOperator setOperator, Double value, XLCFContentType type);
    }
}
