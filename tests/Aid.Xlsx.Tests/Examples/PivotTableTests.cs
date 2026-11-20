using ClosedXML.Examples;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.Examples
{
    [TestFixture]
    public class PivotTableTests
    {
        [Test]
        public void PivotTables()
        {
            TestHelper.RunTestExample<PivotTables>(@"PivotTables\PivotTables.xlsx");
        }
    }
}
