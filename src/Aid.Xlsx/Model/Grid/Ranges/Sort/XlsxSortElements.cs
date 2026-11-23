#nullable disable

using Aid.Xlsx.Infrastructure.Utilities;

namespace Aid.Xlsx.Model.Grid.Ranges.Sort;

internal class XlsxSortElements : IXlsxSortElements
{
    readonly List<IXlsxSortElement> elements = [];
    public void Add(int elementNumber)
    {
        Add(elementNumber, XlsxSortOrder.Ascending);
    }
    public void Add(int elementNumber, XlsxSortOrder sortOrder)
    {
        Add(elementNumber, sortOrder, true);
    }
    public void Add(int elementNumber, XlsxSortOrder sortOrder, bool ignoreBlanks)
    {
        Add(elementNumber, sortOrder, ignoreBlanks, false);
    }
    public void Add(int elementNumber, XlsxSortOrder sortOrder, bool ignoreBlanks, bool matchCase)
    {
        elements.Add(new XlsxSortElement(
            elementNumber,
            sortOrder,
            ignoreBlanks,
            matchCase));
    }

    public void Add(string elementNumber)
    {
        Add(elementNumber, XlsxSortOrder.Ascending);
    }
    public void Add(string elementNumber, XlsxSortOrder sortOrder)
    {
        Add(elementNumber, sortOrder, true);
    }
    public void Add(string elementNumber, XlsxSortOrder sortOrder, bool ignoreBlanks)
    {
        Add(elementNumber, sortOrder, ignoreBlanks, false);
    }
    public void Add(string elementNumber, XlsxSortOrder sortOrder, bool ignoreBlanks, bool matchCase)
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

    public void Remove(int elementNumber)
    {
        elements.RemoveAt(elementNumber - 1);
    }

    internal void AddRange(IEnumerable<XlsxSortElement> sortElements) => elements.AddRange(sortElements);
}
