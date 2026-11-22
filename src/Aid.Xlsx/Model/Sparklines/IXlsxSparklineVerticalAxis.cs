#nullable disable

// Keep this file CodeMaid organised and cleaned
namespace Aid.Xlsx.Sparklines
{
    public interface IXlsxSparklineVerticalAxis
    {
        #region Public Properties

        Double? ManualMax { get; set; }

        Double? ManualMin { get; set; }

        XlsxSparklineAxisMinMax MaxAxisType { get; set; }

        XlsxSparklineAxisMinMax MinAxisType { get; set; }

        #endregion Public Properties

        #region Public Methods

        IXlsxSparklineVerticalAxis SetManualMax(Double? value);

        IXlsxSparklineVerticalAxis SetManualMin(Double? value);

        IXlsxSparklineVerticalAxis SetMaxAxisType(XlsxSparklineAxisMinMax value);

        IXlsxSparklineVerticalAxis SetMinAxisType(XlsxSparklineAxisMinMax value);

        #endregion Public Methods
    }
}
