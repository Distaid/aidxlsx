// Тест, исполняющий пример из отдельного проекта примеров, которого пока нет в репозитории.
// Каталог Examples целиком исключён из компиляции (см. Aid.Xlsx.Tests.csproj).
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