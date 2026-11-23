#nullable disable

// Keep this file CodeMaid organised and cleaned
namespace Aid.Xlsx.Model.Sparklines;

public interface IXlsxSparklineVerticalAxis
{
    #region Public Properties

    double? ManualMax { get; set; }

    double? ManualMin { get; set; }

    XlsxSparklineAxisMinMax MaxAxisType { get; set; }

    XlsxSparklineAxisMinMax MinAxisType { get; set; }

    #endregion Public Properties

    #region Public Methods

    IXlsxSparklineVerticalAxis SetManualMax(double? value);

    IXlsxSparklineVerticalAxis SetManualMin(double? value);

    IXlsxSparklineVerticalAxis SetMaxAxisType(XlsxSparklineAxisMinMax value);

    IXlsxSparklineVerticalAxis SetMinAxisType(XlsxSparklineAxisMinMax value);

    #endregion Public Methods
}
