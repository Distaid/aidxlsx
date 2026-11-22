using Aid.Xlsx.Styles.Formatting;

namespace Aid.Xlsx.Styles
{
    /// <summary>
    /// Reference point of date/number formats available.
    /// See more at: https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat.aspx
    /// </summary>
    public static class XlsxPredefinedFormat
    {
        /// <summary>
        /// General
        /// </summary>
        public static int General { get { return 0; } }

        public enum Number
        {
            /// <summary>
            /// General
            /// </summary>
            General = 0,

            /// <summary>
            /// 0
            /// </summary>
            Integer = 1,

            /// <summary>
            /// 0.00
            /// </summary>
            Precision2 = 2,

            /// <summary>
            /// #,##0
            /// </summary>
            IntegerWithSeparator = 3,

            /// <summary>
            /// #,##0.00
            /// </summary>
            Precision2WithSeparator = 4,

            /// <summary>
            /// 0%
            /// </summary>
            PercentInteger = 9,

            /// <summary>
            /// 0.00%
            /// </summary>
            PercentPrecision2 = 10,

            /// <summary>
            /// 0.00E+00
            /// </summary>
            ScientificPrecision2 = 11,

            /// <summary>
            /// # ?/?
            /// </summary>
            FractionPrecision1 = 12,

            /// <summary>
            /// # ??/??
            /// </summary>
            FractionPrecision2 = 13,

            /// <summary>
            /// #,##0 ,(#,##0)
            /// </summary>
            IntegerWithSeparatorAndParens = 37,

            /// <summary>
            /// #,##0 ,[Red](#,##0)
            /// </summary>
            IntegerWithSeparatorAndParensRed = 38,

            /// <summary>
            /// #,##0.00,(#,##0.00)
            /// </summary>
            Precision2WithSeparatorAndParens = 39,

            /// <summary>
            /// #,##0.00,[Red](#,##0.00)
            /// </summary>
            Precision2WithSeparatorAndParensRed = 40,

            /// <summary>
            /// ##0.0E+0
            /// </summary>
            ScientificUpToHundredsAndPrecision1 = 48,

            /// <summary>
            /// @
            /// </summary>
            Text = 49
        }

        public enum DateTime
        {
            /// <summary>
            /// General
            /// </summary>
            General = 0,

            /// <summary>
            /// d/m/yyyy
            /// </summary>
            DayMonthYear4WithSlashes = 14,

            /// <summary>
            /// d-mmm-yy
            /// </summary>
            DayMonthAbbrYear2WithDashes = 15,

            /// <summary>
            /// d-mmm
            /// </summary>
            DayMonthAbbrWithDash = 16,

            /// <summary>
            /// mmm-yy
            /// </summary>
            MonthAbbrYear2WithDash = 17,

            /// <summary>
            /// h:mm tt
            /// </summary>
            Hour12MinutesAmPm = 18,

            /// <summary>
            /// h:mm:ss tt
            /// </summary>
            Hour12MinutesSecondsAmPm = 19,

            /// <summary>
            /// H:mm
            /// </summary>
            Hour24Minutes = 20,

            /// <summary>
            /// H:mm:ss
            /// </summary>
            Hour24MinutesSeconds = 21,

            /// <summary>
            /// m/d/yyyy H:mm
            /// </summary>
            MonthDayYear4WithDashesHour24Minutes = 22,

            /// <summary>
            /// mm:ss
            /// </summary>
            MinutesSeconds = 45,

            /// <summary>
            /// [h]:mm:ss
            /// </summary>
            Hour12MinutesSeconds = 46,

            /// <summary>
            /// mm:ss.0
            /// </summary>
            /// <remarks>
            /// OOXML specification is missing colon.
            /// </remarks>
            MinutesSecondsMillis1 = 47,

            /// <summary>
            /// @
            /// </summary>
            Text = 49
        }

        internal static IReadOnlyDictionary<int, XlsxNumberFormat> FormatCodes { get; } = new Dictionary<int, XlsxNumberFormat>
        {
            { 0, new XlsxNumberFormat(string.Empty) },
            { 1, new XlsxNumberFormat("0") },
            { 2, new XlsxNumberFormat("0.00") },
            { 3, new XlsxNumberFormat("#,##0") },
            { 4, new XlsxNumberFormat("#,##0.00") },
            { 7, new XlsxNumberFormat("$#,##0.00_);($#,##0.00)") },
            { 9, new XlsxNumberFormat("0%") },
            { 10, new XlsxNumberFormat("0.00%") },
            { 11, new XlsxNumberFormat("0.00E+00") },
            { 12, new XlsxNumberFormat("# ?/?") },
            { 13, new XlsxNumberFormat("# ??/??") },
            { 14, new XlsxNumberFormat("M/d/yyyy") },
            { 15, new XlsxNumberFormat("d-MMM-yy") },
            { 16, new XlsxNumberFormat("d-MMM") },
            { 17, new XlsxNumberFormat("MMM-yy") },
            { 18, new XlsxNumberFormat("h:mm AM/PM") },
            { 19, new XlsxNumberFormat("h:mm:ss AM/PM") },
            { 20, new XlsxNumberFormat("H:mm") },
            { 21, new XlsxNumberFormat("H:mm:ss") },
            { 22, new XlsxNumberFormat("M/d/yyyy H:mm") },
            { 37, new XlsxNumberFormat("#,##0 ;(#,##0)") },
            { 38, new XlsxNumberFormat("#,##0 ;[Red](#,##0)") },
            { 39, new XlsxNumberFormat("#,##0.00;(#,##0.00)") },
            { 40, new XlsxNumberFormat("#,##0.00;[Red](#,##0.00)") },
            { 45, new XlsxNumberFormat("mm:ss") },
            { 46, new XlsxNumberFormat("[h]:mm:ss") },
            { 47, new XlsxNumberFormat("mm:ss.0") },
            { 48, new XlsxNumberFormat("##0.0E+0") },
            { 49, new XlsxNumberFormat("@") },
        };

        internal static IReadOnlyDictionary<XlsxNumberFormat, int> NumberFormatIds { get; } = FormatCodes.ToDictionary(x => x.Value, x => x.Key);
    }
}
