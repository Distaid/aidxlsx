using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.SpreadsheetML.Xml;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.SpreadsheetML.Xml;

internal class SequentialMapTests
{
    [Test]
    public void RetainsReservedIdsAndWritesEntriesInSavedIdOrder()
    {
        var formats = new BiDictionary<int, string>();
        formats.Add(9, "custom");
        formats.Add(8, "percent");
        formats.Add(7, "general");
        var reservedIds = new Dictionary<string, int> { ["percent"] = 9, ["general"] = 0 };

        var map = SequentialMap<int, string>.Create(["custom"], formats, reservedIds, usedStart: 164);

        Assert.That(map.GetSavedId("general"), Is.EqualTo(0));
        Assert.That(map.GetSavedId("percent"), Is.EqualTo(9));
        Assert.That(map.GetSavedId(9), Is.EqualTo(164));
        Assert.That(map.GetActual(), Is.EqualTo(new[] { (0, "general"), (9, "percent"), (164, "custom") }));
    }

    [Test]
    public void ReusesFirstSavedIdForAnActualIdRegisteredTwice()
    {
        var formats = new BiDictionary<int, string>();
        formats.Add(42, "custom");
        var map = new SequentialMap<int, string>(formats);
        map.Add(42);
        map.Add(42);
        map.Sort();

        Assert.That(map.GetSavedId(42), Is.Zero);
        Assert.That(map.GetActual(), Is.EqualTo(new[] { (0, "custom"), (1, "custom") }));
    }
}
