#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Cells;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.Sparklines
{
    public interface IXlsxSparklineGroups : IEnumerable<IXlsxSparklineGroup>
    {
        #region Public Properties

        IXlsxWorksheet Worksheet { get; }

        #endregion Public Properties

        #region Public Methods

        IXlsxSparklineGroup Add(IXlsxSparklineGroup sparklineGroup);
        IXlsxSparklineGroup Add(string locationAddress, string sourceDataAddress);
        IXlsxSparklineGroup Add(IXlsxCell location, IXlsxRange sourceData);
        IXlsxSparklineGroup Add(IXlsxRange locationRange, IXlsxRange sourceDataRange);

        void CopyTo(IXlsxWorksheet targetSheet);

        IXlsxSparkline GetSparkline(IXlsxCell cell);
        IEnumerable<IXlsxSparkline> GetSparklines(IXlsxRangeBase rangeBase);

        void Remove(IXlsxCell cell);
        void Remove(IXlsxRangeBase range);
        void Remove(IXlsxSparklineGroup sparklineGroup);
        void RemoveAll();

        #endregion Public Methods
    }
}
