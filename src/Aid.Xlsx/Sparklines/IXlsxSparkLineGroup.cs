#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Cells;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.Sparklines
{
    public enum XlsxDisplayBlanksAsValues
    {
        Interpolate = 0,
        NotPlotted = 1,
        Zero = 2
    }

    public enum XlsxSparklineAxisMinMax
    {
        Automatic = 0,
        SameForAll = 1,
        Custom = 2
    }

    [Flags]
    public enum XlsxSparklineMarkers
    {
        None = 0,
        HighPoint = 1 << 1,
        LowPoint = 1 << 2,
        FirstPoint = 1 << 3,
        LastPoint = 1 << 4,
        NegativePoints = 1 << 5,
        Markers = 1 << 6,
        All = HighPoint | LowPoint | FirstPoint | LastPoint | NegativePoints | Markers
    }

    public enum XlsxSparklineType
    {
        Line = 0,
        Column = 1,
        Stacked = 2
    }

    public interface IXlsxSparklineGroup : IEnumerable<IXlsxSparkline>
    {
        IXlsxRange DateRange { get; set; }

        XlsxDisplayBlanksAsValues DisplayEmptyCellsAs { get; set; }

        Boolean DisplayHidden { get; set; }

        IXlsxSparklineHorizontalAxis HorizontalAxis { get; }

        Double LineWeight { get; set; }

        XlsxSparklineMarkers ShowMarkers { get; set; }

        IXlsxSparklineStyle Style { get; set; }

        XlsxSparklineType Type { get; set; }

        IXlsxSparklineVerticalAxis VerticalAxis { get; }

        IXlsxWorksheet Worksheet { get; }

        IXlsxSparkline Add(IXlsxCell location, IXlsxRange sourceData);

        IEnumerable<IXlsxSparkline> Add(IXlsxRange locationRange, IXlsxRange sourceDataRange);

        IEnumerable<IXlsxSparkline> Add(string locationRangeAddress, string sourceDataAddress);

        void CopyFrom(IXlsxSparklineGroup sparklineGroup);

        /// <summary>
        /// Copy this sparkline group to the specified worksheet
        /// </summary>
        /// <param name="targetSheet">The worksheet to copy this sparkline group to</param>
        IXlsxSparklineGroup CopyTo(IXlsxWorksheet targetSheet);

        IXlsxSparkline GetSparkline(IXlsxCell cell);

        IEnumerable<IXlsxSparkline> GetSparklines(IXlsxRangeBase searchRange);

        void Remove(IXlsxCell cell);

        void Remove(IXlsxSparkline sparkline);

        void RemoveAll();

        IXlsxSparklineGroup SetDateRange(IXlsxRange value);

        IXlsxSparklineGroup SetDisplayEmptyCellsAs(XlsxDisplayBlanksAsValues value);

        IXlsxSparklineGroup SetDisplayHidden(Boolean value);

        IXlsxSparklineGroup SetLineWeight(Double value);

        IXlsxSparklineGroup SetShowMarkers(XlsxSparklineMarkers value);

        IXlsxSparklineGroup SetStyle(IXlsxSparklineStyle value);

        IXlsxSparklineGroup SetType(XlsxSparklineType value);
    }
}
