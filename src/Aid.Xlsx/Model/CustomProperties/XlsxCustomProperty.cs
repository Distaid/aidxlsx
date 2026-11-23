#nullable disable

using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Model.CustomProperties;

internal class XlsxCustomProperty : IXlsxCustomProperty
{
    private readonly XlsxWorkbook _workbook;

    private String name;

    public XlsxCustomProperty(XlsxWorkbook workbook)
    {
        _workbook = workbook;
    }

    #region IXlsxCustomProperty Members

    public String Name
    {
        get { return name; }
        set
        {
            if (name == value) return;

            if (_workbook.CustomProperties.Any(t => t.Name == value))
                throw new ArgumentException(
                    String.Format("This workbook already contains a custom property named '{0}'", value));

            name = value;
        }
    }

    public XlsxCustomPropertyType Type
    {
        get
        {
            if (Value is DateTime)
                return XlsxCustomPropertyType.Date;

            if (Value is Boolean)
                return XlsxCustomPropertyType.Boolean;
            if (Double.TryParse(Value.ToString(), out _))
                return XlsxCustomPropertyType.Number;

            return XlsxCustomPropertyType.Text;
        }
    }

    public Object Value { get; set; }

    public T GetValue<T>()
    {
        return (T)Convert.ChangeType(Value, typeof(T));
    }

    #endregion
}
