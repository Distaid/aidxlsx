#nullable disable

namespace Aid.Xlsx.Model.CustomProperties;

public enum XlsxCustomPropertyType { Text, Number, Date, Boolean }
public interface IXlsxCustomProperty
{
    String Name { get; set; }
    XlsxCustomPropertyType Type { get; }
    Object Value { get; set; }
    T GetValue<T>();
}
