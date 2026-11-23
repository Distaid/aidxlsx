// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.DataImport.RowReaders;

internal class NullDataReader : IInsertDataReader
{
    private readonly XlsxCellValue[] _row = [Blank.Value];
    private readonly int _count;

    public NullDataReader(IEnumerable<object> nulls)
    {
        _count = nulls.Count();
    }

    public IEnumerable<IEnumerable<XlsxCellValue>> GetRecords()
    {
        return Enumerable.Repeat(_row, _count);
    }

    public int GetPropertiesCount()
    {
        return 0;
    }

    public string? GetPropertyName(int propertyIndex)
    {
        return null;
    }
}
