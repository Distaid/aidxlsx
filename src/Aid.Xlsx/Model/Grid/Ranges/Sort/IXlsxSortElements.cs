#nullable disable

namespace Aid.Xlsx.Model.Grid.Ranges.Sort;

public interface IXlsxSortElements : IEnumerable<IXlsxSortElement>
{
    void Add(int elementNumber);
    void Add(int elementNumber, XlsxSortOrder sortOrder);
    void Add(int elementNumber, XlsxSortOrder sortOrder, bool ignoreBlanks);
    void Add(int elementNumber, XlsxSortOrder sortOrder, bool ignoreBlanks, bool matchCase);

    void Add(string elementNumber);
    void Add(string elementNumber, XlsxSortOrder sortOrder);
    void Add(string elementNumber, XlsxSortOrder sortOrder, bool ignoreBlanks);
    void Add(string elementNumber, XlsxSortOrder sortOrder, bool ignoreBlanks, bool matchCase);

    void Clear();

    void Remove(int elementNumber);
}
