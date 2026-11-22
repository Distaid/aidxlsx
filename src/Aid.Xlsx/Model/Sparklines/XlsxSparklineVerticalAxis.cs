#nullable disable

// Keep this file CodeMaid organised and cleaned
namespace Aid.Xlsx.Sparklines
{
    internal class XlsxSparklineVerticalAxis : IXlsxSparklineVerticalAxis
    {
        #region Public Properties

        public Double? ManualMax
        {
            get => _manualMax;
            set => SetManualMax(value);
        }

        public Double? ManualMin
        {
            get => _manualMin;
            set => SetManualMin(value);
        }

        public XlsxSparklineAxisMinMax MaxAxisType
        {
            get => _maxAxisType;
            set => SetMaxAxisType(value);
        }

        public XlsxSparklineAxisMinMax MinAxisType
        {
            get => _minAxisType;
            set => SetMinAxisType(value);
        }

        #endregion Public Properties

        #region Public Methods

        public IXlsxSparklineVerticalAxis SetManualMax(Double? manualMax)
        {
            if (manualMax != null)
                MaxAxisType = XlsxSparklineAxisMinMax.Custom;

            _manualMax = manualMax;
            return this;
        }

        public IXlsxSparklineVerticalAxis SetManualMin(Double? manualMin)
        {
            if (manualMin != null)
                MinAxisType = XlsxSparklineAxisMinMax.Custom;

            _manualMin = manualMin;
            return this;
        }

        public IXlsxSparklineVerticalAxis SetMaxAxisType(XlsxSparklineAxisMinMax maxAxisType)
        {
            if (maxAxisType != XlsxSparklineAxisMinMax.Custom)
                _manualMax = null;

            _maxAxisType = maxAxisType;
            return this;
        }

        public IXlsxSparklineVerticalAxis SetMinAxisType(XlsxSparklineAxisMinMax minAxisType)
        {
            if (minAxisType != XlsxSparklineAxisMinMax.Custom)
                _manualMin = null;

            _minAxisType = minAxisType;
            return this;
        }

        #endregion Public Methods

        #region Private Fields

        private Double? _manualMax;
        private Double? _manualMin;
        private XlsxSparklineAxisMinMax _maxAxisType;
        private XlsxSparklineAxisMinMax _minAxisType;

        #endregion Private Fields

        public IXlsxSparklineGroup SparklineGroup { get; }

        public XlsxSparklineVerticalAxis(IXlsxSparklineGroup sparklineGroup)
        {
            SparklineGroup = sparklineGroup ?? throw new ArgumentNullException(nameof(sparklineGroup));
        }

        public static void Copy(IXlsxSparklineVerticalAxis from, IXlsxSparklineVerticalAxis to)
        {
            to.ManualMax = from.ManualMax;
            to.ManualMin = from.ManualMin;
            to.MaxAxisType = from.MaxAxisType;
            to.MinAxisType = from.MinAxisType;
        }
    }
}
