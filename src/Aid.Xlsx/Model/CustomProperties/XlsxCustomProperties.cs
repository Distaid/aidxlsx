#nullable disable

using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Model.CustomProperties;

internal class XlsxCustomProperties(XlsxWorkbook workbook) : IXlsxCustomProperties, IEnumerable<IXlsxCustomProperty>
{
    readonly XlsxWorkbook workbook = workbook;
    private readonly Dictionary<String, IXlsxCustomProperty> customProperties = [];
    public void Add(IXlsxCustomProperty customProperty)
    {
        customProperties.Add(customProperty.Name, customProperty);
    }
    public void Add<T>(String name, T value)
    {
        var cp = new XlsxCustomProperty(workbook) { Name = name, Value = value };
        Add(cp);
    }

    public void Delete(String name)
    {
        customProperties.Remove(name);
    }
    public IXlsxCustomProperty CustomProperty(String name)
    {
        return customProperties[name];
    }

    public IEnumerator<IXlsxCustomProperty> GetEnumerator()
    {
        return customProperties.Values.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }


}
