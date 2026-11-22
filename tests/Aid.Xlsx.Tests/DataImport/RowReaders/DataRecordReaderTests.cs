using System.Data;
using NUnit.Framework;
using Aid.Xlsx.DataImport.RowReaders;

namespace Aid.Xlsx.Tests.DataImport.RowReaders;

public class DataRecordReaderTests
{
    private static IEnumerable<IDataRecord> GetData()
    {
        using var table = new DataTable();
        table.Columns.Add("StringValue", typeof(string));
        table.Columns.Add("NumericValue", typeof(int));
        table.Rows.Add("Value 1", 100);
        table.Rows.Add("Value 2", DBNull.Value);
        table.Rows.Add("Value 3", 300);

        using var reader = table.CreateDataReader();
        while (reader.Read())
            yield return reader;
    }

    [Test]
    public void CanGetPropertyName()
    {
        var reader = InsertDataReaderFactory.Instance.CreateReader(GetData());
        Assert.AreEqual("StringValue", reader.GetPropertyName(0));
        Assert.AreEqual("NumericValue", reader.GetPropertyName(1));
    }

    [Test]
    public void CanGetPropertiesCount()
    {
        var reader = InsertDataReaderFactory.Instance.CreateReader(GetData());
        Assert.AreEqual(2, reader.GetPropertiesCount());
    }

    [Test]
    public void CanGetRecordsCount()
    {
        var reader = InsertDataReaderFactory.Instance.CreateReader(GetData());
        Assert.AreEqual(3, reader.GetRecords().Count());
    }

    [Test]
    public void CanGetData()
    {
        var reader = InsertDataReaderFactory.Instance.CreateReader(GetData());
        var result = reader.GetRecords().ToArray();

        Assert.AreEqual("Value 1", result.First().First());
        Assert.AreEqual(100, result.First().Last());
        Assert.That(result[1].Last().GetText(), Is.Empty);
        Assert.AreEqual("Value 3", result.Last().First());
        Assert.AreEqual(300, result.Last().Last());
    }
}
