// A test that executes an example from the separate examples project, which is not in this repository yet.
// The whole Examples directory is excluded from compilation (see Aid.Xlsx.Tests.csproj).
using ClosedXML.Examples.Sparklines;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.Excel.Sparklines
{
    public partial class SparklinesTests
    {
        #region Test Examples

        [Test]
        public void CreateSampleSparklines()
        {
            TestHelper.RunTestExample<SampleSparklines>(@"Sparklines\SampleSparklines.xlsx");
        }

        #endregion Test Examples
    }
}
