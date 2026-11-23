#nullable disable

using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Model.CustomProperties;

internal class XlsxCustomProperty(XlsxWorkbook workbook) : IXlsxCustomProperty
{
    private readonly XlsxWorkbook _workbook = workbook;

    private string name;

    #region IXlsxCustomProperty Members

    public string Name
    {
        get { return name; }
        set
        {
            if (name == value) return;

            if (_workbook.CustomProperties.Any(t => t.Name == value))
                throw new ArgumentException(
                    string.Format("This workbook already contains a custom property named '{0}'", value));

            name = value;
        }
    }

    public XlsxCustomPropertyType Type
    {
        get
        {
            if (Value is DateTime)
                return XlsxCustomPropertyType.Date;

            if (Value is bool)
                return XlsxCustomPropertyType.Boolean;
            if (double.TryParse(Value.ToString(), out _))
                return XlsxCustomPropertyType.Number;

            return XlsxCustomPropertyType.Text;
        }
    }

    public object Value { get; set; }

    public T GetValue<T>()
    {
        return (T)Convert.ChangeType(Value, typeof(T));
    }

    #endregion
}
