// Keep this file CodeMaid organised and cleaned
using System.Collections;
using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.DataImport.RowReaders;

internal class ArrayReader : IInsertDataReader
{
    private readonly IEnumerable<IEnumerable> _data;

    public ArrayReader(IEnumerable<IEnumerable> data)
    {
        _data = data ?? throw new ArgumentNullException(nameof(data));
    }

    public IEnumerable<IEnumerable<XlsxCellValue>> GetRecords()
    {
        return _data.Select(item => item.Cast<object>().Select(XlsxCellValue.FromInsertedObject));
    }

    public int GetPropertiesCount()
    {
        if (!_data.Any())
            return 0;

        return _data.First().Cast<object>().Count();
    }

    public string? GetPropertyName(int propertyIndex)
    {
        return null;
    }
}
