#nullable disable

using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.PageSetup;

internal class XlsxPrintAreas : IXlsxPrintAreas
{
    readonly List<IXlsxRange> ranges = [];
    private readonly XlsxWorksheet worksheet;
    public XlsxPrintAreas(XlsxWorksheet worksheet)
    {
        this.worksheet = worksheet;
    }

    public XlsxPrintAreas(XlsxPrintAreas defaultPrintAreas, XlsxWorksheet worksheet)
    {
        ranges = [.. defaultPrintAreas.ranges];
        this.worksheet = worksheet;
    }

    public void Clear()
    {
        ranges.Clear();
    }

    public void Add(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
    {
        ranges.Add(worksheet.Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn));
    }

    public void Add(string rangeAddress)
    {
        ranges.Add(worksheet.Range(rangeAddress));
    }

    public void Add(string firstCellAddress, string lastCellAddress)
    {
        ranges.Add(worksheet.Range(firstCellAddress, lastCellAddress));
    }

    public void Add(IXlsxAddress firstCellAddress, IXlsxAddress lastCellAddress)
    {
        ranges.Add(worksheet.Range(firstCellAddress, lastCellAddress));
    }

    public IEnumerator<IXlsxRange> GetEnumerator()
    {
        return ranges.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
