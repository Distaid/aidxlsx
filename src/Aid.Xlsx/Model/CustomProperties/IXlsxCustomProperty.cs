#nullable disable

namespace Aid.Xlsx.Model.CustomProperties;

public enum XlsxCustomPropertyType { Text, Number, Date, Boolean }
public interface IXlsxCustomProperty
{
    string Name { get; set; }
    XlsxCustomPropertyType Type { get; }
    object Value { get; set; }
    T GetValue<T>();
}
