#nullable disable

using Aid.Xlsx.Workbooks;

namespace Aid.Xlsx.CustomProperties
{
    internal class XlsxCustomProperties : IXlsxCustomProperties, IEnumerable<IXlsxCustomProperty>
    {
        XlsxWorkbook workbook;
        public XlsxCustomProperties(XlsxWorkbook workbook)
        {
            this.workbook = workbook;
        }

        private Dictionary<String, IXlsxCustomProperty> customProperties = new Dictionary<String, IXlsxCustomProperty>();
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
}
