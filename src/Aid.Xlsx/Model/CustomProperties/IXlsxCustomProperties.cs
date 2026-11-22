#nullable disable

namespace Aid.Xlsx.Model.CustomProperties;

public interface IXlsxCustomProperties : IEnumerable<IXlsxCustomProperty>
{
    void Add(IXlsxCustomProperty customProperty);
    void Add<T>(String name, T value);
    void Delete(String name);
    IXlsxCustomProperty CustomProperty(String name);

}
