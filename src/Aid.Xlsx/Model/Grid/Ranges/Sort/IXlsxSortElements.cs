#nullable disable

namespace Aid.Xlsx.Model.Grid.Ranges.Sort;

public interface IXlsxSortElements : IEnumerable<IXlsxSortElement>
{
    void Add(Int32 elementNumber);
    void Add(Int32 elementNumber, XlsxSortOrder sortOrder);
    void Add(Int32 elementNumber, XlsxSortOrder sortOrder, Boolean ignoreBlanks);
    void Add(Int32 elementNumber, XlsxSortOrder sortOrder, Boolean ignoreBlanks, Boolean matchCase);

    void Add(String elementNumber);
    void Add(String elementNumber, XlsxSortOrder sortOrder);
    void Add(String elementNumber, XlsxSortOrder sortOrder, Boolean ignoreBlanks);
    void Add(String elementNumber, XlsxSortOrder sortOrder, Boolean ignoreBlanks, Boolean matchCase);

    void Clear();

    void Remove(Int32 elementNumber);
}
