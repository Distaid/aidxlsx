#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Styles.Colors;

namespace Aid.Xlsx.Sparklines
{
    public interface IXlsxSparklineHorizontalAxis
    {
        #region Public Properties

        XlsxColor Color { get; set; }

        Boolean DateAxis { get; }

        Boolean IsVisible { get; set; }

        Boolean RightToLeft { get; set; }

        #endregion Public Properties

        #region Public Methods

        IXlsxSparklineHorizontalAxis SetColor(XlsxColor value);

        IXlsxSparklineHorizontalAxis SetRightToLeft(Boolean value);

        IXlsxSparklineHorizontalAxis SetVisible(Boolean value);

        #endregion Public Methods
    }
}
