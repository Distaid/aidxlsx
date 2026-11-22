#nullable disable

using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.PageSetup
{
    internal class XlsxPrintAreas : IXlsxPrintAreas
    {
        List<IXlsxRange> ranges = new List<IXlsxRange>();
        private XlsxWorksheet worksheet;
        public XlsxPrintAreas(XlsxWorksheet worksheet)
        {
            this.worksheet = worksheet;
        }

        public XlsxPrintAreas(XlsxPrintAreas defaultPrintAreas, XlsxWorksheet worksheet)
        {
            ranges = defaultPrintAreas.ranges.ToList();
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
}
