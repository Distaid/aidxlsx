#nullable disable

// Keep this file CodeMaid organised and cleaned
using System.Data;
using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.DataImport.RowReaders;

internal class DataRecordReader : IInsertDataReader
{
    private readonly IEnumerable<XlsxCellValue>[] _inMemoryData;
    private string[] _columns;

    public DataRecordReader(IEnumerable<IDataRecord> data)
    {
        ArgumentNullException.ThrowIfNull(data);

        _inMemoryData = [.. ReadToEnd(data)];
    }

    public IEnumerable<IEnumerable<XlsxCellValue>> GetRecords()
    {
        return _inMemoryData;
    }

    public int GetPropertiesCount()
    {
        return _columns.Length;
    }

    public string GetPropertyName(int propertyIndex)
    {
        if (propertyIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(propertyIndex), "Property index must be non-negative");

        if (_columns == null)
            return null;

        if (propertyIndex >= _columns.Length)
            throw new ArgumentOutOfRangeException($"{propertyIndex} exceeds the number of the table columns");

        return _columns[propertyIndex];
    }

    private IEnumerable<IEnumerable<XlsxCellValue>> ReadToEnd(IEnumerable<IDataRecord> data)
    {
        foreach (IDataRecord dataRecord in data)
        {
            yield return ToEnumerable(dataRecord).ToArray();
        }
    }

    private IEnumerable<XlsxCellValue> ToEnumerable(IDataRecord dataRecord)
    {
        var firstRow = false;
        if (_columns == null)
        {
            firstRow = true;
            _columns = new string[dataRecord.FieldCount];
        }

        for (int i = 0; i < dataRecord.FieldCount; i++)
        {
            if (firstRow)
                _columns[i] = dataRecord.GetName(i);

            var value = dataRecord[i];
            yield return XlsxCellValue.FromInsertedObject(value);
        }
    }
}
