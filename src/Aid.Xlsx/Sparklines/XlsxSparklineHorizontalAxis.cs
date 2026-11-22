#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Styles.Colors;

namespace Aid.Xlsx.Sparklines
{
    internal class XlsxSparklineHorizontalAxis : IXlsxSparklineHorizontalAxis
    {
        #region Public Properties

        public XlsxColor Color { get; set; }

        public bool DateAxis => SparklineGroup.DateRange != null;

        public bool IsVisible { get; set; }

        public bool RightToLeft { get; set; }

        public IXlsxSparklineGroup SparklineGroup { get; }

        #endregion Public Properties

        #region Public Constructors

        public XlsxSparklineHorizontalAxis(IXlsxSparklineGroup sparklineGroup)
        {
            SparklineGroup = sparklineGroup ?? throw new ArgumentNullException(nameof(sparklineGroup));
        }

        #endregion Public Constructors

        #region Public Methods

        public IXlsxSparklineHorizontalAxis SetColor(XlsxColor value)
        {
            Color = value ?? throw new ArgumentNullException(nameof(value));
            return this;
        }

        public IXlsxSparklineHorizontalAxis SetRightToLeft(bool value)
        {
            RightToLeft = value;
            return this;
        }

        public IXlsxSparklineHorizontalAxis SetVisible(bool value)
        {
            IsVisible = value;
            return this;
        }

        #endregion Public Methods

        public static void Copy(IXlsxSparklineHorizontalAxis from, IXlsxSparklineHorizontalAxis to)
        {
            to.Color = from.Color;
            to.IsVisible = from.IsVisible;
            to.RightToLeft = from.RightToLeft;
        }
    }
}
