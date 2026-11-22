// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Internal.IO;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Styles;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Aid.Xlsx.Utils
{
    internal static class OpenXmlHelper
    {
        #region Public Methods

        /// <summary>
        /// Convert color in Aid.Xlsx representation to specified OpenXML type.
        /// </summary>
        /// <typeparam name="T">The descendant of <see cref="ColorType"/>.</typeparam>
        /// <param name="openXMLColor">The existing instance of ColorType.</param>
        /// <param name="xlColor">Color in Aid.Xlsx format.</param>
        /// <param name="isDifferential">Flag specifying that the color should be saved in
        /// differential format (affects the transparent color processing).</param>
        /// <returns>The original color in OpenXML format.</returns>
        public static T FromXlsxColor<T>(this ColorType openXMLColor, XlsxColor xlColor, bool isDifferential = false)
            where T : ColorType
        {
            var adapter = new ColorTypeAdapter(openXMLColor);
            FillFromXlsxColor(adapter, xlColor, isDifferential);
            return (T)adapter.ColorType;
        }

        /// <summary>
        /// Convert color in Aid.Xlsx representation to specified OpenXML type.
        /// </summary>
        /// <typeparam name="T">The descendant of <see cref="X14.ColorType"/>.</typeparam>
        /// <param name="openXMLColor">The existing instance of ColorType.</param>
        /// <param name="xlColor">Color in Aid.Xlsx format.</param>
        /// <param name="isDifferential">Flag specifying that the color should be saved in
        /// differential format (affects the transparent color processing).</param>
        /// <returns>The original color in OpenXML format.</returns>
        public static T FromXlsxColor<T>(this X14.ColorType openXMLColor, XlsxColor xlColor, bool isDifferential = false)
            where T : X14.ColorType
        {
            var adapter = new X14ColorTypeAdapter(openXMLColor);
            FillFromXlsxColor(adapter, xlColor, isDifferential);
            return (T)adapter.ColorType;
        }

        public static BooleanValue? GetBooleanValue(bool value, bool? defaultValue = null)
        {
            return (defaultValue.HasValue && value == defaultValue.Value) ? null : new BooleanValue(value);
        }

        public static bool GetBooleanValueAsBool(BooleanValue? value, bool defaultValue)
        {
            return (value?.HasValue ?? false) ? value.Value : defaultValue;
        }

        /// <summary>
        /// Convert color in OpenXML representation to the Aid.Xlsx color type.
        /// </summary>
        /// <param name="openXMLColor">Color in OpenXML format.</param>
        /// <returns>The color in Aid.Xlsx format.</returns>
        public static XlsxColor ToXlsxColor(this ColorType openXMLColor)
        {
            return ConvertToXlsxColor(new ColorTypeAdapter(openXMLColor));
        }

        /// <summary>
        /// Convert color in OpenXML representation to the Aid.Xlsx color type.
        /// </summary>
        /// <param name="openXMLColor">Color in OpenXML format.</param>
        /// <returns>The color in Aid.Xlsx format.</returns>
        public static XlsxColor ToXlsxColor(this X14.ColorType openXMLColor)
        {
            return ConvertToXlsxColor(new X14ColorTypeAdapter(openXMLColor));
        }

#nullable disable

        internal static void LoadNumberFormat(NumberingFormat nfSource, IXlsxNumberFormat nf)
        {
            if (nfSource == null) return;

            if (nfSource.NumberFormatId != null && nfSource.NumberFormatId.Value < XlsxConstants.NumberOfBuiltInStyles)
                nf.NumberFormatId = (Int32)nfSource.NumberFormatId.Value;
            else if (nfSource.FormatCode != null)
                nf.Format = nfSource.FormatCode.Value;
        }

        internal static void LoadBorder(Border borderSource, IXlsxBorder border)
        {
            if (borderSource == null) return;

            LoadBorderValues(borderSource.DiagonalBorder, border.SetDiagonalBorder, border.SetDiagonalBorderColor);

            if (borderSource.DiagonalUp != null)
                border.DiagonalUp = borderSource.DiagonalUp.Value;
            if (borderSource.DiagonalDown != null)
                border.DiagonalDown = borderSource.DiagonalDown.Value;

            LoadBorderValues(borderSource.LeftBorder, border.SetLeftBorder, border.SetLeftBorderColor);
            LoadBorderValues(borderSource.RightBorder, border.SetRightBorder, border.SetRightBorderColor);
            LoadBorderValues(borderSource.TopBorder, border.SetTopBorder, border.SetTopBorderColor);
            LoadBorderValues(borderSource.BottomBorder, border.SetBottomBorder, border.SetBottomBorderColor);
        }

        private static void LoadBorderValues(BorderPropertiesType source, Func<XlsxBorderStyleValues, IXlsxStyle> setBorder, Func<XlsxColor, IXlsxStyle> setColor)
        {
            if (source != null)
            {
                if (source.Style != null)
                    setBorder(source.Style.Value.ToXlsx());
                if (source.Color != null)
                    setColor(source.Color.ToXlsxColor());
            }
        }

        // Differential fills store the patterns differently than other fills
        // Actually differential fills make more sense. bg is bg and fg is fg
        // 'Other' fills store the bg color in the fg field when pattern type is solid
        internal static void LoadFill(Fill openXMLFill, IXlsxFill xlsxFill, Boolean differentialFillFormat)
        {
            if (openXMLFill == null || openXMLFill.PatternFill == null) return;

            if (openXMLFill.PatternFill.PatternType != null)
                xlsxFill.PatternType = openXMLFill.PatternFill.PatternType.Value.ToXlsx();
            else
                xlsxFill.PatternType = XlsxFillPatternValues.Solid;

            switch (xlsxFill.PatternType)
            {
                case XlsxFillPatternValues.None:
                    break;

                case XlsxFillPatternValues.Solid:
                    if (differentialFillFormat)
                    {
                        if (openXMLFill.PatternFill.BackgroundColor != null)
                            xlsxFill.BackgroundColor = openXMLFill.PatternFill.BackgroundColor.ToXlsxColor();
                        else
                            xlsxFill.BackgroundColor = XlsxColor.FromIndex(64);
                    }
                    else
                    {
                        // yes, source is foreground!
                        if (openXMLFill.PatternFill.ForegroundColor != null)
                            xlsxFill.BackgroundColor = openXMLFill.PatternFill.ForegroundColor.ToXlsxColor();
                        else
                            xlsxFill.BackgroundColor = XlsxColor.FromIndex(64);
                    }
                    break;

                default:
                    if (openXMLFill.PatternFill.ForegroundColor != null)
                        xlsxFill.PatternColor = openXMLFill.PatternFill.ForegroundColor.ToXlsxColor();

                    if (openXMLFill.PatternFill.BackgroundColor != null)
                        xlsxFill.BackgroundColor = openXMLFill.PatternFill.BackgroundColor.ToXlsxColor();
                    else
                        xlsxFill.BackgroundColor = XlsxColor.FromIndex(64);
                    break;
            }
        }

#nullable enable

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Here we perform the actual conversion from OpenXML color to Aid.Xlsx color.
        /// </summary>
        /// <param name="openXMLColor">OpenXML color. Must be either <see cref="ColorType"/> or <see cref="X14.ColorType"/>.
        /// Since these types do not implement a common interface we use dynamic.</param>
        /// <returns>The color in Aid.Xlsx format.</returns>
        private static XlsxColor ConvertToXlsxColor(IColorTypeAdapter openXMLColor)
        {
            XlsxColor? retVal = null;
            if (openXMLColor.Rgb?.Value is not null)
            {
                var thisColor = ColorStringParser.ParseFromArgb(openXMLColor.Rgb.Value.AsSpan());
                retVal = XlsxColor.FromColor(thisColor);
            }
            else if (openXMLColor.Indexed is not null && openXMLColor.Indexed <= 64)
                retVal = XlsxColor.FromIndex((Int32)openXMLColor.Indexed.Value);
            else if (openXMLColor.Theme is not null)
            {
                retVal = openXMLColor.Tint is not null
                    ? XlsxColor.FromTheme((XlsxThemeColor)openXMLColor.Theme.Value, openXMLColor.Tint.Value)
                    : XlsxColor.FromTheme((XlsxThemeColor)openXMLColor.Theme.Value);
            }
            return retVal ?? XlsxColor.Automatic;
        }

        /// <summary>
        /// Initialize properties of the existing instance of the color in OpenXML format basing on properties of the color
        /// in Aid.Xlsx format.
        /// </summary>
        /// <param name="openXMLColor">OpenXML color. Must be either <see cref="ColorType"/> or <see cref="X14.ColorType"/>.
        /// Since these types do not implement a common interface we use dynamic.</param>
        /// <param name="xlColor">Color in Aid.Xlsx format.</param>
        /// <param name="isDifferential">Flag specifying that the color should be saved in
        /// differential format (affects the transparent color processing).</param>
        private static void FillFromXlsxColor(IColorTypeAdapter openXMLColor, XlsxColor xlColor, bool isDifferential)
        {
            if (openXMLColor == null)
                throw new ArgumentNullException(nameof(openXMLColor));

            if (xlColor == null)
                throw new ArgumentNullException(nameof(xlColor));

            switch (xlColor.ColorType)
            {
                case XlsxColorType.Color:
                    openXMLColor.Rgb = xlColor.Color.ToHex();
                    break;

                case XlsxColorType.Indexed:
                    // 64 is 'transparent' and should be ignored for differential formats
                    if (!isDifferential || xlColor.Indexed != 64)
                        openXMLColor.Indexed = (UInt32)xlColor.Indexed;
                    break;

                case XlsxColorType.Theme:
                    openXMLColor.Theme = (UInt32)xlColor.ThemeColor;

                    if (xlColor.ThemeTint != 0)
                        openXMLColor.Tint = xlColor.ThemeTint;
                    break;
            }
        }

        internal static int NormalizeRotation(uint textRotation)
        {
            return textRotation switch
            {
                <= 90 => (int)textRotation,
                <= 180 => 90 - (int)textRotation,
                255 => 255,
                _ => throw new ArgumentOutOfRangeException()
            };
        }

        #endregion Private Methods
    }
}
