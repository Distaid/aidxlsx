#nullable disable

// Keep this file CodeMaid organised and cleaned

using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Sparklines;

public interface IXlsxSparklineStyle
{
    #region Public Properties

    XlsxColor FirstMarkerColor { get; set; }

    XlsxColor HighMarkerColor { get; set; }

    XlsxColor LastMarkerColor { get; set; }

    XlsxColor LowMarkerColor { get; set; }

    XlsxColor MarkersColor { get; set; }

    XlsxColor NegativeColor { get; set; }

    XlsxColor SeriesColor { get; set; }

    #endregion Public Properties

    #region Public Methods

    IXlsxSparklineStyle SetFirstMarkerColor(XlsxColor value);

    IXlsxSparklineStyle SetHighMarkerColor(XlsxColor value);

    IXlsxSparklineStyle SetLastMarkerColor(XlsxColor value);

    IXlsxSparklineStyle SetLowMarkerColor(XlsxColor value);

    IXlsxSparklineStyle SetMarkersColor(XlsxColor value);

    IXlsxSparklineStyle SetNegativeColor(XlsxColor value);

    IXlsxSparklineStyle SetSeriesColor(XlsxColor value);

    #endregion Public Methods
}
