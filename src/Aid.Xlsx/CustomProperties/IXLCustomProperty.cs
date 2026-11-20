#nullable disable

using System;

namespace Aid.Xlsx
{
    public enum XLCustomPropertyType { Text, Number, Date, Boolean }
    public interface IXLCustomProperty
    {
        String Name { get; set; }
        XLCustomPropertyType Type { get; }
        Object Value { get; set; }
        T GetValue<T>();
    }
}
