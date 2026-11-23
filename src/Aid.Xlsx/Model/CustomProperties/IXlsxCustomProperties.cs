#nullable disable

namespace Aid.Xlsx.Model.CustomProperties;

public interface IXlsxCustomProperties : IEnumerable<IXlsxCustomProperty>
{
    void Add(IXlsxCustomProperty customProperty);
    void Add<T>(string name, T value);
    void Delete(string name);
    IXlsxCustomProperty CustomProperty(string name);

}
