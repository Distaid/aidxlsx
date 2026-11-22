#nullable disable

using Aid.Xlsx.Infrastructure.Utilities;

namespace Aid.Xlsx.Model.Grid.Ranges.Sort;

internal class XlsxSortElements : IXlsxSortElements
{
    List<IXlsxSortElement> elements = new List<IXlsxSortElement>();
    public void Add(Int32 elementNumber)
    {
        Add(elementNumber, XlsxSortOrder.Ascending);
    }
    public void Add(Int32 elementNumber, XlsxSortOrder sortOrder)
    {
        Add(elementNumber, sortOrder, true);
    }
    public void Add(Int32 elementNumber, XlsxSortOrder sortOrder, Boolean ignoreBlanks)
    {
        Add(elementNumber, sortOrder, ignoreBlanks, false);
    }
    public void Add(Int32 elementNumber, XlsxSortOrder sortOrder, Boolean ignoreBlanks, Boolean matchCase)
    {
        elements.Add(new XlsxSortElement(
            elementNumber,
            sortOrder,
            ignoreBlanks,
            matchCase));
    }

    public void Add(String elementNumber)
    {
        Add(elementNumber, XlsxSortOrder.Ascending);
    }
    public void Add(String elementNumber, XlsxSortOrder sortOrder)
    {
        Add(elementNumber, sortOrder, true);
    }
    public void Add(String elementNumber, XlsxSortOrder sortOrder, Boolean ignoreBlanks)
    {
        Add(elementNumber, sortOrder, ignoreBlanks, false);
    }
    public void Add(String elementNumber, XlsxSortOrder sortOrder, Boolean ignoreBlanks, Boolean matchCase)
    {
        elements.Add(new XlsxSortElement(
            XlsxHelper.GetColumnNumberFromLetter(elementNumber),
            sortOrder,
            ignoreBlanks,
            matchCase));
    }

    public IEnumerator<IXlsxSortElement> GetEnumerator()
    {
        return elements.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public void Clear()
    {
        elements.Clear();
    }

    public void Remove(Int32 elementNumber)
    {
        elements.RemoveAt(elementNumber - 1);
    }

    internal void AddRange(IEnumerable<XlsxSortElement> sortElements) => elements.AddRange(sortElements);
}
