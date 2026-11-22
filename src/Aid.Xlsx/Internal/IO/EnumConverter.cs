#nullable disable

using Aid.Xlsx.ConditionalFormats;
using Aid.Xlsx.DataValidation;
using Aid.Xlsx.Drawings;
using Aid.Xlsx.PageSetup;
using Aid.Xlsx.PivotTables.Areas;
using Aid.Xlsx.PivotTables.PivotValues;
using Aid.Xlsx.PivotTables;
using Aid.Xlsx.RichText;
using Aid.Xlsx.Sparklines;
using Aid.Xlsx.Styles;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Aid.Xlsx.Internal.IO
{
    internal static class EnumConverter
    {
        #region To OpenXml

        public static UnderlineValues ToOpenXml(this XlsxFontUnderlineValues value)
        {
            switch (value)
            {
                case XlsxFontUnderlineValues.Double:
                    return UnderlineValues.Double;

                case XlsxFontUnderlineValues.DoubleAccounting:
                    return UnderlineValues.DoubleAccounting;

                case XlsxFontUnderlineValues.None:
                    return UnderlineValues.None;

                case XlsxFontUnderlineValues.Single:
                    return UnderlineValues.Single;

                case XlsxFontUnderlineValues.SingleAccounting:
                    return UnderlineValues.SingleAccounting;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        private static readonly String[] XlsxFontUnderlineValuesStrings =
        {
            "double",
            "doubleAccounting",
            "none",
            "single",
            "singleAccounting"
        };

        public static string ToOpenXmlString(this XlsxFontUnderlineValues value)
            => XlsxFontUnderlineValuesStrings[(int)value];

        public static OrientationValues ToOpenXml(this XlsxPageOrientation value)
        {
            switch (value)
            {
                case XlsxPageOrientation.Default:
                    return OrientationValues.Default;

                case XlsxPageOrientation.Landscape:
                    return OrientationValues.Landscape;

                case XlsxPageOrientation.Portrait:
                    return OrientationValues.Portrait;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static VerticalAlignmentRunValues ToOpenXml(this XlsxFontVerticalTextAlignmentValues value)
        {
            switch (value)
            {
                case XlsxFontVerticalTextAlignmentValues.Baseline:
                    return VerticalAlignmentRunValues.Baseline;

                case XlsxFontVerticalTextAlignmentValues.Subscript:
                    return VerticalAlignmentRunValues.Subscript;

                case XlsxFontVerticalTextAlignmentValues.Superscript:
                    return VerticalAlignmentRunValues.Superscript;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        private static readonly String[] XlsxFontVerticalTextAlignmentValuesStrings =
        {
            "baseline",
            "subscript",
            "superscript"
        };

        public static String ToOpenXmlString(this XlsxFontVerticalTextAlignmentValues value)
            => XlsxFontVerticalTextAlignmentValuesStrings[(int)value];

        private static readonly String[] XlsxFontSchemeStrings =
        {
            "none",
            "major",
            "minor"
        };

        public static String ToOpenXml(this XlsxFontScheme value)
            => XlsxFontSchemeStrings[(int)value];

        public static FontSchemeValues ToOpenXmlEnum(this XlsxFontScheme value)
        {
            return value switch
            {
                XlsxFontScheme.None => FontSchemeValues.None,
                XlsxFontScheme.Major => FontSchemeValues.Major,
                XlsxFontScheme.Minor => FontSchemeValues.Minor,
                _ => throw new ArgumentOutOfRangeException()
            };
        }

        public static PatternValues ToOpenXml(this XlsxFillPatternValues value)
        {
            switch (value)
            {
                case XlsxFillPatternValues.DarkDown:
                    return PatternValues.DarkDown;

                case XlsxFillPatternValues.DarkGray:
                    return PatternValues.DarkGray;

                case XlsxFillPatternValues.DarkGrid:
                    return PatternValues.DarkGrid;

                case XlsxFillPatternValues.DarkHorizontal:
                    return PatternValues.DarkHorizontal;

                case XlsxFillPatternValues.DarkTrellis:
                    return PatternValues.DarkTrellis;

                case XlsxFillPatternValues.DarkUp:
                    return PatternValues.DarkUp;

                case XlsxFillPatternValues.DarkVertical:
                    return PatternValues.DarkVertical;

                case XlsxFillPatternValues.Gray0625:
                    return PatternValues.Gray0625;

                case XlsxFillPatternValues.Gray125:
                    return PatternValues.Gray125;

                case XlsxFillPatternValues.LightDown:
                    return PatternValues.LightDown;

                case XlsxFillPatternValues.LightGray:
                    return PatternValues.LightGray;

                case XlsxFillPatternValues.LightGrid:
                    return PatternValues.LightGrid;

                case XlsxFillPatternValues.LightHorizontal:
                    return PatternValues.LightHorizontal;

                case XlsxFillPatternValues.LightTrellis:
                    return PatternValues.LightTrellis;

                case XlsxFillPatternValues.LightUp:
                    return PatternValues.LightUp;

                case XlsxFillPatternValues.LightVertical:
                    return PatternValues.LightVertical;

                case XlsxFillPatternValues.MediumGray:
                    return PatternValues.MediumGray;

                case XlsxFillPatternValues.None:
                    return PatternValues.None;

                case XlsxFillPatternValues.Solid:
                    return PatternValues.Solid;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static BorderStyleValues ToOpenXml(this XlsxBorderStyleValues value)
        {
            switch (value)
            {
                case XlsxBorderStyleValues.DashDot:
                    return BorderStyleValues.DashDot;

                case XlsxBorderStyleValues.DashDotDot:
                    return BorderStyleValues.DashDotDot;

                case XlsxBorderStyleValues.Dashed:
                    return BorderStyleValues.Dashed;

                case XlsxBorderStyleValues.Dotted:
                    return BorderStyleValues.Dotted;

                case XlsxBorderStyleValues.Double:
                    return BorderStyleValues.Double;

                case XlsxBorderStyleValues.Hair:
                    return BorderStyleValues.Hair;

                case XlsxBorderStyleValues.Medium:
                    return BorderStyleValues.Medium;

                case XlsxBorderStyleValues.MediumDashDot:
                    return BorderStyleValues.MediumDashDot;

                case XlsxBorderStyleValues.MediumDashDotDot:
                    return BorderStyleValues.MediumDashDotDot;

                case XlsxBorderStyleValues.MediumDashed:
                    return BorderStyleValues.MediumDashed;

                case XlsxBorderStyleValues.None:
                    return BorderStyleValues.None;

                case XlsxBorderStyleValues.SlantDashDot:
                    return BorderStyleValues.SlantDashDot;

                case XlsxBorderStyleValues.Thick:
                    return BorderStyleValues.Thick;

                case XlsxBorderStyleValues.Thin:
                    return BorderStyleValues.Thin;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static HorizontalAlignmentValues ToOpenXml(this XlsxAlignmentHorizontalValues value)
        {
            switch (value)
            {
                case XlsxAlignmentHorizontalValues.Center:
                    return HorizontalAlignmentValues.Center;

                case XlsxAlignmentHorizontalValues.CenterContinuous:
                    return HorizontalAlignmentValues.CenterContinuous;

                case XlsxAlignmentHorizontalValues.Distributed:
                    return HorizontalAlignmentValues.Distributed;

                case XlsxAlignmentHorizontalValues.Fill:
                    return HorizontalAlignmentValues.Fill;

                case XlsxAlignmentHorizontalValues.General:
                    return HorizontalAlignmentValues.General;

                case XlsxAlignmentHorizontalValues.Justify:
                    return HorizontalAlignmentValues.Justify;

                case XlsxAlignmentHorizontalValues.Left:
                    return HorizontalAlignmentValues.Left;

                case XlsxAlignmentHorizontalValues.Right:
                    return HorizontalAlignmentValues.Right;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static VerticalAlignmentValues ToOpenXml(this XlsxAlignmentVerticalValues value)
        {
            switch (value)
            {
                case XlsxAlignmentVerticalValues.Bottom:
                    return VerticalAlignmentValues.Bottom;

                case XlsxAlignmentVerticalValues.Center:
                    return VerticalAlignmentValues.Center;

                case XlsxAlignmentVerticalValues.Distributed:
                    return VerticalAlignmentValues.Distributed;

                case XlsxAlignmentVerticalValues.Justify:
                    return VerticalAlignmentValues.Justify;

                case XlsxAlignmentVerticalValues.Top:
                    return VerticalAlignmentValues.Top;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static PageOrderValues ToOpenXml(this XlsxPageOrderValues value)
        {
            switch (value)
            {
                case XlsxPageOrderValues.DownThenOver:
                    return PageOrderValues.DownThenOver;

                case XlsxPageOrderValues.OverThenDown:
                    return PageOrderValues.OverThenDown;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static CellCommentsValues ToOpenXml(this XlsxShowCommentsValues value)
        {
            switch (value)
            {
                case XlsxShowCommentsValues.AsDisplayed:
                    return CellCommentsValues.AsDisplayed;

                case XlsxShowCommentsValues.AtEnd:
                    return CellCommentsValues.AtEnd;

                case XlsxShowCommentsValues.None:
                    return CellCommentsValues.None;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static PrintErrorValues ToOpenXml(this XlsxPrintErrorValues value)
        {
            switch (value)
            {
                case XlsxPrintErrorValues.Blank:
                    return PrintErrorValues.Blank;

                case XlsxPrintErrorValues.Dash:
                    return PrintErrorValues.Dash;

                case XlsxPrintErrorValues.Displayed:
                    return PrintErrorValues.Displayed;

                case XlsxPrintErrorValues.NA:
                    return PrintErrorValues.NA;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static CalculateModeValues ToOpenXml(this XlsxCalculateMode value)
        {
            switch (value)
            {
                case XlsxCalculateMode.Auto:
                    return CalculateModeValues.Auto;

                case XlsxCalculateMode.AutoNoTable:
                    return CalculateModeValues.AutoNoTable;

                case XlsxCalculateMode.Manual:
                    return CalculateModeValues.Manual;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static ReferenceModeValues ToOpenXml(this XlsxReferenceStyle value)
        {
            switch (value)
            {
                case XlsxReferenceStyle.R1C1:
                    return ReferenceModeValues.R1C1;

                case XlsxReferenceStyle.A1:
                    return ReferenceModeValues.A1;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static uint ToOpenXml(this XlsxAlignmentReadingOrderValues value)
        {
            switch (value)
            {
                case XlsxAlignmentReadingOrderValues.ContextDependent:
                    return 0;

                case XlsxAlignmentReadingOrderValues.LeftToRight:
                    return 1;

                case XlsxAlignmentReadingOrderValues.RightToLeft:
                    return 2;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static DataValidationValues ToOpenXml(this XlsxAllowedValues value)
        {
            switch (value)
            {
                case XlsxAllowedValues.AnyValue:
                    return DataValidationValues.None;

                case XlsxAllowedValues.Custom:
                    return DataValidationValues.Custom;

                case XlsxAllowedValues.Date:
                    return DataValidationValues.Date;

                case XlsxAllowedValues.Decimal:
                    return DataValidationValues.Decimal;

                case XlsxAllowedValues.List:
                    return DataValidationValues.List;

                case XlsxAllowedValues.TextLength:
                    return DataValidationValues.TextLength;

                case XlsxAllowedValues.Time:
                    return DataValidationValues.Time;

                case XlsxAllowedValues.WholeNumber:
                    return DataValidationValues.Whole;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static DataValidationErrorStyleValues ToOpenXml(this XlsxErrorStyle value)
        {
            switch (value)
            {
                case XlsxErrorStyle.Information:
                    return DataValidationErrorStyleValues.Information;

                case XlsxErrorStyle.Warning:
                    return DataValidationErrorStyleValues.Warning;

                case XlsxErrorStyle.Stop:
                    return DataValidationErrorStyleValues.Stop;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static DataValidationOperatorValues ToOpenXml(this XlsxOperator value)
        {
            switch (value)
            {
                case XlsxOperator.Between:
                    return DataValidationOperatorValues.Between;

                case XlsxOperator.EqualOrGreaterThan:
                    return DataValidationOperatorValues.GreaterThanOrEqual;

                case XlsxOperator.EqualOrLessThan:
                    return DataValidationOperatorValues.LessThanOrEqual;

                case XlsxOperator.EqualTo:
                    return DataValidationOperatorValues.Equal;

                case XlsxOperator.GreaterThan:
                    return DataValidationOperatorValues.GreaterThan;

                case XlsxOperator.LessThan:
                    return DataValidationOperatorValues.LessThan;

                case XlsxOperator.NotBetween:
                    return DataValidationOperatorValues.NotBetween;

                case XlsxOperator.NotEqualTo:
                    return DataValidationOperatorValues.NotEqual;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static SheetStateValues ToOpenXml(this XlsxWorksheetVisibility value)
        {
            switch (value)
            {
                case XlsxWorksheetVisibility.Visible:
                    return SheetStateValues.Visible;

                case XlsxWorksheetVisibility.Hidden:
                    return SheetStateValues.Hidden;

                case XlsxWorksheetVisibility.VeryHidden:
                    return SheetStateValues.VeryHidden;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        private static readonly String[] XlsxPhoneticAlignmentStrings =
        {
            "center",
            "distributed",
            "left",
            "noControl"
        };

        public static String ToOpenXmlString(this XlsxPhoneticAlignment value)
            => XlsxPhoneticAlignmentStrings[(int)value];

        public static PhoneticValues ToOpenXml(this XlsxPhoneticType value)
        {
            switch (value)
            {
                case XlsxPhoneticType.FullWidthKatakana:
                    return PhoneticValues.FullWidthKatakana;

                case XlsxPhoneticType.HalfWidthKatakana:
                    return PhoneticValues.HalfWidthKatakana;

                case XlsxPhoneticType.Hiragana:
                    return PhoneticValues.Hiragana;

                case XlsxPhoneticType.NoConversion:
                    return PhoneticValues.NoConversion;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        private static readonly String[] XlsxPhoneticTypeStrings =
        {
            "fullwidthKatakana",
            "halfwidthKatakana",
            "Hiragana",
            "noConversion"
        };

        public static String ToOpenXmlString(this XlsxPhoneticType value)
            => XlsxPhoneticTypeStrings[(int)value];

        public static DataConsolidateFunctionValues ToOpenXml(this XlsxPivotSummary value)
        {
            switch (value)
            {
                case XlsxPivotSummary.Sum: return DataConsolidateFunctionValues.Sum;
                case XlsxPivotSummary.Count: return DataConsolidateFunctionValues.Count;
                case XlsxPivotSummary.Average: return DataConsolidateFunctionValues.Average;
                case XlsxPivotSummary.Minimum: return DataConsolidateFunctionValues.Minimum;
                case XlsxPivotSummary.Maximum: return DataConsolidateFunctionValues.Maximum;
                case XlsxPivotSummary.Product: return DataConsolidateFunctionValues.Product;
                case XlsxPivotSummary.CountNumbers: return DataConsolidateFunctionValues.CountNumbers;
                case XlsxPivotSummary.StandardDeviation: return DataConsolidateFunctionValues.StandardDeviation;
                case XlsxPivotSummary.PopulationStandardDeviation: return DataConsolidateFunctionValues.StandardDeviationP;
                case XlsxPivotSummary.Variance: return DataConsolidateFunctionValues.Variance;
                case XlsxPivotSummary.PopulationVariance: return DataConsolidateFunctionValues.VarianceP;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static ShowDataAsValues ToOpenXml(this XlsxPivotCalculation value)
        {
            switch (value)
            {
                case XlsxPivotCalculation.Normal: return ShowDataAsValues.Normal;
                case XlsxPivotCalculation.DifferenceFrom: return ShowDataAsValues.Difference;
                case XlsxPivotCalculation.PercentageOf: return ShowDataAsValues.Percent;
                case XlsxPivotCalculation.PercentageDifferenceFrom: return ShowDataAsValues.PercentageDifference;
                case XlsxPivotCalculation.RunningTotal: return ShowDataAsValues.RunTotal;
                case XlsxPivotCalculation.PercentageOfRow: return ShowDataAsValues.PercentOfRaw; // There's a typo in the OpenXML SDK =)
                case XlsxPivotCalculation.PercentageOfColumn: return ShowDataAsValues.PercentOfColumn;
                case XlsxPivotCalculation.PercentageOfTotal: return ShowDataAsValues.PercentOfTotal;
                case XlsxPivotCalculation.Index: return ShowDataAsValues.Index;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static SheetViewValues ToOpenXml(this XlsxSheetViewOptions value)
        {
            switch (value)
            {
                case XlsxSheetViewOptions.Normal: return SheetViewValues.Normal;
                case XlsxSheetViewOptions.PageBreakPreview: return SheetViewValues.PageBreakPreview;
                case XlsxSheetViewOptions.PageLayout: return SheetViewValues.PageLayout;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static ConditionalFormatValues ToOpenXml(this XlsxConditionalFormatType value)
        {
            switch (value)
            {
                case XlsxConditionalFormatType.Expression: return ConditionalFormatValues.Expression;
                case XlsxConditionalFormatType.CellIs: return ConditionalFormatValues.CellIs;
                case XlsxConditionalFormatType.ColorScale: return ConditionalFormatValues.ColorScale;
                case XlsxConditionalFormatType.DataBar: return ConditionalFormatValues.DataBar;
                case XlsxConditionalFormatType.IconSet: return ConditionalFormatValues.IconSet;
                case XlsxConditionalFormatType.Top10: return ConditionalFormatValues.Top10;
                case XlsxConditionalFormatType.IsUnique: return ConditionalFormatValues.UniqueValues;
                case XlsxConditionalFormatType.IsDuplicate: return ConditionalFormatValues.DuplicateValues;
                case XlsxConditionalFormatType.ContainsText: return ConditionalFormatValues.ContainsText;
                case XlsxConditionalFormatType.NotContainsText: return ConditionalFormatValues.NotContainsText;
                case XlsxConditionalFormatType.StartsWith: return ConditionalFormatValues.BeginsWith;
                case XlsxConditionalFormatType.EndsWith: return ConditionalFormatValues.EndsWith;
                case XlsxConditionalFormatType.IsBlank: return ConditionalFormatValues.ContainsBlanks;
                case XlsxConditionalFormatType.NotBlank: return ConditionalFormatValues.NotContainsBlanks;
                case XlsxConditionalFormatType.IsError: return ConditionalFormatValues.ContainsErrors;
                case XlsxConditionalFormatType.NotError: return ConditionalFormatValues.NotContainsErrors;
                case XlsxConditionalFormatType.TimePeriod: return ConditionalFormatValues.TimePeriod;
                case XlsxConditionalFormatType.AboveAverage: return ConditionalFormatValues.AboveAverage;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static ConditionalFormatValueObjectValues ToOpenXml(this XlsxCFContentType value)
        {
            switch (value)
            {
                case XlsxCFContentType.Number: return ConditionalFormatValueObjectValues.Number;
                case XlsxCFContentType.Percent: return ConditionalFormatValueObjectValues.Percent;
                case XlsxCFContentType.Maximum: return ConditionalFormatValueObjectValues.Max;
                case XlsxCFContentType.Minimum: return ConditionalFormatValueObjectValues.Min;
                case XlsxCFContentType.Formula: return ConditionalFormatValueObjectValues.Formula;
                case XlsxCFContentType.Percentile: return ConditionalFormatValueObjectValues.Percentile;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static ConditionalFormattingOperatorValues ToOpenXml(this XlsxCFOperator value)
        {
            switch (value)
            {
                case XlsxCFOperator.LessThan: return ConditionalFormattingOperatorValues.LessThan;
                case XlsxCFOperator.EqualOrLessThan: return ConditionalFormattingOperatorValues.LessThanOrEqual;
                case XlsxCFOperator.Equal: return ConditionalFormattingOperatorValues.Equal;
                case XlsxCFOperator.NotEqual: return ConditionalFormattingOperatorValues.NotEqual;
                case XlsxCFOperator.EqualOrGreaterThan: return ConditionalFormattingOperatorValues.GreaterThanOrEqual;
                case XlsxCFOperator.GreaterThan: return ConditionalFormattingOperatorValues.GreaterThan;
                case XlsxCFOperator.Between: return ConditionalFormattingOperatorValues.Between;
                case XlsxCFOperator.NotBetween: return ConditionalFormattingOperatorValues.NotBetween;
                case XlsxCFOperator.Contains: return ConditionalFormattingOperatorValues.ContainsText;
                case XlsxCFOperator.NotContains: return ConditionalFormattingOperatorValues.NotContains;
                case XlsxCFOperator.StartsWith: return ConditionalFormattingOperatorValues.BeginsWith;
                case XlsxCFOperator.EndsWith: return ConditionalFormattingOperatorValues.EndsWith;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static IconSetValues ToOpenXml(this XlsxIconSetStyle value)
        {
            switch (value)
            {
                case XlsxIconSetStyle.ThreeArrows: return IconSetValues.ThreeArrows;
                case XlsxIconSetStyle.ThreeArrowsGray: return IconSetValues.ThreeArrowsGray;
                case XlsxIconSetStyle.ThreeFlags: return IconSetValues.ThreeFlags;
                case XlsxIconSetStyle.ThreeTrafficLights1: return IconSetValues.ThreeTrafficLights1;
                case XlsxIconSetStyle.ThreeTrafficLights2: return IconSetValues.ThreeTrafficLights2;
                case XlsxIconSetStyle.ThreeSigns: return IconSetValues.ThreeSigns;
                case XlsxIconSetStyle.ThreeSymbols: return IconSetValues.ThreeSymbols;
                case XlsxIconSetStyle.ThreeSymbols2: return IconSetValues.ThreeSymbols2;
                case XlsxIconSetStyle.FourArrows: return IconSetValues.FourArrows;
                case XlsxIconSetStyle.FourArrowsGray: return IconSetValues.FourArrowsGray;
                case XlsxIconSetStyle.FourRedToBlack: return IconSetValues.FourRedToBlack;
                case XlsxIconSetStyle.FourRating: return IconSetValues.FourRating;
                case XlsxIconSetStyle.FourTrafficLights: return IconSetValues.FourTrafficLights;
                case XlsxIconSetStyle.FiveArrows: return IconSetValues.FiveArrows;
                case XlsxIconSetStyle.FiveArrowsGray: return IconSetValues.FiveArrowsGray;
                case XlsxIconSetStyle.FiveRating: return IconSetValues.FiveRating;
                case XlsxIconSetStyle.FiveQuarters: return IconSetValues.FiveQuarters;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static TimePeriodValues ToOpenXml(this XlsxTimePeriod value)
        {
            switch (value)
            {
                case XlsxTimePeriod.Yesterday: return TimePeriodValues.Yesterday;
                case XlsxTimePeriod.Today: return TimePeriodValues.Today;
                case XlsxTimePeriod.Tomorrow: return TimePeriodValues.Tomorrow;
                case XlsxTimePeriod.InTheLast7Days: return TimePeriodValues.Last7Days;
                case XlsxTimePeriod.LastWeek: return TimePeriodValues.LastWeek;
                case XlsxTimePeriod.ThisWeek: return TimePeriodValues.ThisWeek;
                case XlsxTimePeriod.NextWeek: return TimePeriodValues.NextWeek;
                case XlsxTimePeriod.LastMonth: return TimePeriodValues.LastMonth;
                case XlsxTimePeriod.ThisMonth: return TimePeriodValues.ThisMonth;
                case XlsxTimePeriod.NextMonth: return TimePeriodValues.NextMonth;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        private static readonly IReadOnlyDictionary<XlsxPictureFormat, PartTypeInfo> PictureFormatMap =
            new Dictionary<XlsxPictureFormat, PartTypeInfo>
            {
                { XlsxPictureFormat.Unknown, new PartTypeInfo("image/unknown", ".bin") },
                { XlsxPictureFormat.Bmp, ImagePartType.Bmp },
                { XlsxPictureFormat.Gif, ImagePartType.Gif },
                { XlsxPictureFormat.Png, ImagePartType.Png },
                { XlsxPictureFormat.Tiff, ImagePartType.Tiff },
                { XlsxPictureFormat.Icon, ImagePartType.Icon },
                { XlsxPictureFormat.Pcx, ImagePartType.Pcx },
                { XlsxPictureFormat.Jpeg, ImagePartType.Jpeg },
                { XlsxPictureFormat.Emf, ImagePartType.Emf },
                { XlsxPictureFormat.Wmf, ImagePartType.Wmf },
                { XlsxPictureFormat.Webp, new PartTypeInfo("image/webp", ".webp") }
            };

        public static PartTypeInfo ToOpenXml(this XlsxPictureFormat value)
        {
            return PictureFormatMap[value];
        }

        public static Xdr.EditAsValues ToOpenXml(this XlsxPicturePlacement value)
        {
            switch (value)
            {
                case XlsxPicturePlacement.FreeFloating:
                    return Xdr.EditAsValues.Absolute;

                case XlsxPicturePlacement.Move:
                    return Xdr.EditAsValues.OneCell;

                case XlsxPicturePlacement.MoveAndSize:
                    return Xdr.EditAsValues.TwoCell;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static PivotAreaValues ToOpenXml(this XlsxPivotAreaType value)
        {
            switch (value)
            {
                case XlsxPivotAreaType.None:
                    return PivotAreaValues.None;

                case XlsxPivotAreaType.Normal:
                    return PivotAreaValues.Normal;

                case XlsxPivotAreaType.Data:
                    return PivotAreaValues.Data;

                case XlsxPivotAreaType.All:
                    return PivotAreaValues.All;

                case XlsxPivotAreaType.Origin:
                    return PivotAreaValues.Origin;

                case XlsxPivotAreaType.Button:
                    return PivotAreaValues.Button;

                case XlsxPivotAreaType.TopRight:
                    return PivotAreaValues.TopRight;

                case XlsxPivotAreaType.TopEnd:
                    return PivotAreaValues.TopEnd;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "XlsxPivotAreaValues value not implemented");
            }
        }

        public static X14.SparklineTypeValues ToOpenXml(this XlsxSparklineType value)
        {
            switch (value)
            {
                case XlsxSparklineType.Line: return X14.SparklineTypeValues.Line;
                case XlsxSparklineType.Column: return X14.SparklineTypeValues.Column;
                case XlsxSparklineType.Stacked: return X14.SparklineTypeValues.Stacked;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static X14.SparklineAxisMinMaxValues ToOpenXml(this XlsxSparklineAxisMinMax value)
        {
            switch (value)
            {
                case XlsxSparklineAxisMinMax.Automatic: return X14.SparklineAxisMinMaxValues.Individual;
                case XlsxSparklineAxisMinMax.SameForAll: return X14.SparklineAxisMinMaxValues.Group;
                case XlsxSparklineAxisMinMax.Custom: return X14.SparklineAxisMinMaxValues.Custom;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static X14.DisplayBlanksAsValues ToOpenXml(this XlsxDisplayBlanksAsValues value)
        {
            switch (value)
            {
                case XlsxDisplayBlanksAsValues.Interpolate: return X14.DisplayBlanksAsValues.Span;
                case XlsxDisplayBlanksAsValues.NotPlotted: return X14.DisplayBlanksAsValues.Gap;
                case XlsxDisplayBlanksAsValues.Zero: return X14.DisplayBlanksAsValues.Zero;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        public static FieldSortValues ToOpenXml(this XlsxPivotSortType value)
        {
            switch (value)
            {
                case XlsxPivotSortType.Default: return FieldSortValues.Manual;
                case XlsxPivotSortType.Ascending: return FieldSortValues.Ascending;
                case XlsxPivotSortType.Descending: return FieldSortValues.Descending;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        #endregion To OpenXml

        #region To Xlsx

        private static readonly IReadOnlyDictionary<UnderlineValues, XlsxFontUnderlineValues> UnderlineValuesMap =
            new Dictionary<UnderlineValues, XlsxFontUnderlineValues>
            {
                { UnderlineValues.Double, XlsxFontUnderlineValues.Double },
                { UnderlineValues.DoubleAccounting, XlsxFontUnderlineValues.DoubleAccounting },
                { UnderlineValues.None, XlsxFontUnderlineValues.None },
                { UnderlineValues.Single, XlsxFontUnderlineValues.Single },
                { UnderlineValues.SingleAccounting, XlsxFontUnderlineValues.SingleAccounting },
            };

        public static XlsxFontUnderlineValues ToXlsx(this UnderlineValues value)
        {
            return UnderlineValuesMap[value];
        }

        private static readonly IReadOnlyDictionary<FontSchemeValues, XlsxFontScheme> FontSchemeMap =
            new Dictionary<FontSchemeValues, XlsxFontScheme>
            {
                { FontSchemeValues.None, XlsxFontScheme.None },
                { FontSchemeValues.Major, XlsxFontScheme.Major },
                { FontSchemeValues.Minor, XlsxFontScheme.Minor },
            };

        public static XlsxFontScheme ToXlsx(this FontSchemeValues value)
        {
            return FontSchemeMap[value];
        }

        private static readonly IReadOnlyDictionary<OrientationValues, XlsxPageOrientation> OrientationMap =
            new Dictionary<OrientationValues, XlsxPageOrientation>
            {
                { OrientationValues.Default, XlsxPageOrientation.Default },
                { OrientationValues.Landscape, XlsxPageOrientation.Landscape },
                { OrientationValues.Portrait, XlsxPageOrientation.Portrait },
            };

        public static XlsxPageOrientation ToXlsx(this OrientationValues value)
        {
            return OrientationMap[value];
        }

        private static readonly IReadOnlyDictionary<VerticalAlignmentRunValues, XlsxFontVerticalTextAlignmentValues> VerticalAlignmentRunMap =
            new Dictionary<VerticalAlignmentRunValues, XlsxFontVerticalTextAlignmentValues>
            {
                { VerticalAlignmentRunValues.Baseline, XlsxFontVerticalTextAlignmentValues.Baseline },
                { VerticalAlignmentRunValues.Subscript, XlsxFontVerticalTextAlignmentValues.Subscript },
                { VerticalAlignmentRunValues.Superscript, XlsxFontVerticalTextAlignmentValues.Superscript },
            };


        public static XlsxFontVerticalTextAlignmentValues ToXlsx(this VerticalAlignmentRunValues value)
        {
            return VerticalAlignmentRunMap[value];
        }

        private static readonly IReadOnlyDictionary<PatternValues, XlsxFillPatternValues> PatternMap =
            new Dictionary<PatternValues, XlsxFillPatternValues>
            {
                { PatternValues.DarkDown, XlsxFillPatternValues.DarkDown },
                { PatternValues.DarkGray, XlsxFillPatternValues.DarkGray },
                { PatternValues.DarkGrid, XlsxFillPatternValues.DarkGrid },
                { PatternValues.DarkHorizontal, XlsxFillPatternValues.DarkHorizontal },
                { PatternValues.DarkTrellis, XlsxFillPatternValues.DarkTrellis },
                { PatternValues.DarkUp, XlsxFillPatternValues.DarkUp },
                { PatternValues.DarkVertical, XlsxFillPatternValues.DarkVertical },
                { PatternValues.Gray0625, XlsxFillPatternValues.Gray0625 },
                { PatternValues.Gray125, XlsxFillPatternValues.Gray125 },
                { PatternValues.LightDown, XlsxFillPatternValues.LightDown },
                { PatternValues.LightGray, XlsxFillPatternValues.LightGray },
                { PatternValues.LightGrid, XlsxFillPatternValues.LightGrid },
                { PatternValues.LightHorizontal, XlsxFillPatternValues.LightHorizontal },
                { PatternValues.LightTrellis, XlsxFillPatternValues.LightTrellis },
                { PatternValues.LightUp, XlsxFillPatternValues.LightUp },
                { PatternValues.LightVertical, XlsxFillPatternValues.LightVertical },
                { PatternValues.MediumGray, XlsxFillPatternValues.MediumGray },
                { PatternValues.None, XlsxFillPatternValues.None },
                { PatternValues.Solid, XlsxFillPatternValues.Solid },
            };

        public static XlsxFillPatternValues ToXlsx(this PatternValues value)
        {
            return PatternMap[value];
        }

        private static readonly IReadOnlyDictionary<BorderStyleValues, XlsxBorderStyleValues> BorderStyleMap =
            new Dictionary<BorderStyleValues, XlsxBorderStyleValues>
            {
                { BorderStyleValues.DashDot, XlsxBorderStyleValues.DashDot },
                { BorderStyleValues.DashDotDot, XlsxBorderStyleValues.DashDotDot },
                { BorderStyleValues.Dashed, XlsxBorderStyleValues.Dashed },
                { BorderStyleValues.Dotted, XlsxBorderStyleValues.Dotted },
                { BorderStyleValues.Double, XlsxBorderStyleValues.Double },
                { BorderStyleValues.Hair, XlsxBorderStyleValues.Hair },
                { BorderStyleValues.Medium, XlsxBorderStyleValues.Medium },
                { BorderStyleValues.MediumDashDot, XlsxBorderStyleValues.MediumDashDot },
                { BorderStyleValues.MediumDashDotDot, XlsxBorderStyleValues.MediumDashDotDot },
                { BorderStyleValues.MediumDashed, XlsxBorderStyleValues.MediumDashed },
                { BorderStyleValues.None, XlsxBorderStyleValues.None },
                { BorderStyleValues.SlantDashDot, XlsxBorderStyleValues.SlantDashDot },
                { BorderStyleValues.Thick, XlsxBorderStyleValues.Thick },
                { BorderStyleValues.Thin, XlsxBorderStyleValues.Thin },
            };

        public static XlsxBorderStyleValues ToXlsx(this BorderStyleValues value)
        {
            return BorderStyleMap[value];
        }

        private static readonly IReadOnlyDictionary<HorizontalAlignmentValues, XlsxAlignmentHorizontalValues> HorizontalAlignmentMap =
            new Dictionary<HorizontalAlignmentValues, XlsxAlignmentHorizontalValues>
            {
                { HorizontalAlignmentValues.Center, XlsxAlignmentHorizontalValues.Center },
                { HorizontalAlignmentValues.CenterContinuous, XlsxAlignmentHorizontalValues.CenterContinuous },
                { HorizontalAlignmentValues.Distributed, XlsxAlignmentHorizontalValues.Distributed },
                { HorizontalAlignmentValues.Fill, XlsxAlignmentHorizontalValues.Fill },
                { HorizontalAlignmentValues.General, XlsxAlignmentHorizontalValues.General },
                { HorizontalAlignmentValues.Justify, XlsxAlignmentHorizontalValues.Justify },
                { HorizontalAlignmentValues.Left, XlsxAlignmentHorizontalValues.Left },
                { HorizontalAlignmentValues.Right, XlsxAlignmentHorizontalValues.Right },
            };

        public static XlsxAlignmentHorizontalValues ToXlsx(this HorizontalAlignmentValues value)
        {
            return HorizontalAlignmentMap[value];
        }

        private static readonly IReadOnlyDictionary<VerticalAlignmentValues, XlsxAlignmentVerticalValues> VerticalAlignmentMap =
            new Dictionary<VerticalAlignmentValues, XlsxAlignmentVerticalValues>
            {
                { VerticalAlignmentValues.Bottom, XlsxAlignmentVerticalValues.Bottom },
                { VerticalAlignmentValues.Center, XlsxAlignmentVerticalValues.Center },
                { VerticalAlignmentValues.Distributed, XlsxAlignmentVerticalValues.Distributed },
                { VerticalAlignmentValues.Justify, XlsxAlignmentVerticalValues.Justify },
                { VerticalAlignmentValues.Top, XlsxAlignmentVerticalValues.Top },
            };

        public static XlsxAlignmentVerticalValues ToXlsx(this VerticalAlignmentValues value)
        {
            return VerticalAlignmentMap[value];
        }

        private static readonly IReadOnlyDictionary<PageOrderValues, XlsxPageOrderValues> PageOrdersMap =
            new Dictionary<PageOrderValues, XlsxPageOrderValues>
            {
                { PageOrderValues.DownThenOver, XlsxPageOrderValues.DownThenOver },
                { PageOrderValues.OverThenDown, XlsxPageOrderValues.OverThenDown },
            };

        public static XlsxPageOrderValues ToXlsx(this PageOrderValues value)
        {
            return PageOrdersMap[value];
        }

        private static readonly IReadOnlyDictionary<CellCommentsValues, XlsxShowCommentsValues> CellCommentsMap =
            new Dictionary<CellCommentsValues, XlsxShowCommentsValues>
            {
                { CellCommentsValues.AsDisplayed, XlsxShowCommentsValues.AsDisplayed },
                { CellCommentsValues.AtEnd, XlsxShowCommentsValues.AtEnd },
                { CellCommentsValues.None, XlsxShowCommentsValues.None },
            };

        public static XlsxShowCommentsValues ToXlsx(this CellCommentsValues value)
        {
            return CellCommentsMap[value];
        }

        private static readonly IReadOnlyDictionary<PrintErrorValues, XlsxPrintErrorValues> PrintErrorMap =
            new Dictionary<PrintErrorValues, XlsxPrintErrorValues>
            {
                { PrintErrorValues.Blank, XlsxPrintErrorValues.Blank },
                { PrintErrorValues.Dash, XlsxPrintErrorValues.Dash },
                { PrintErrorValues.Displayed, XlsxPrintErrorValues.Displayed },
                { PrintErrorValues.NA, XlsxPrintErrorValues.NA },
            };

        public static XlsxPrintErrorValues ToXlsx(this PrintErrorValues value)
        {
            return PrintErrorMap[value];
        }

        private static readonly IReadOnlyDictionary<CalculateModeValues, XlsxCalculateMode> CalculateModeMap =
            new Dictionary<CalculateModeValues, XlsxCalculateMode>
            {
                { CalculateModeValues.Auto, XlsxCalculateMode.Auto },
                { CalculateModeValues.AutoNoTable, XlsxCalculateMode.AutoNoTable },
                { CalculateModeValues.Manual, XlsxCalculateMode.Manual },
            };

        public static XlsxCalculateMode ToXlsx(this CalculateModeValues value)
        {
            return CalculateModeMap[value];
        }

        private static readonly IReadOnlyDictionary<ReferenceModeValues, XlsxReferenceStyle> ReferenceModeMap =
            new Dictionary<ReferenceModeValues, XlsxReferenceStyle>
            {
                { ReferenceModeValues.R1C1, XlsxReferenceStyle.R1C1 },
                { ReferenceModeValues.A1, XlsxReferenceStyle.A1 },
            };

        public static XlsxReferenceStyle ToXlsx(this ReferenceModeValues value)
        {
            return ReferenceModeMap[value];
        }

        public static XlsxAlignmentReadingOrderValues ToXlsx(this uint value)
        {
            switch (value)
            {
                case 0:
                    return XlsxAlignmentReadingOrderValues.ContextDependent;

                case 1:
                    return XlsxAlignmentReadingOrderValues.LeftToRight;

                case 2:
                    return XlsxAlignmentReadingOrderValues.RightToLeft;

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), "Not implemented value!");
            }
        }

        private static readonly IReadOnlyDictionary<DataValidationValues, XlsxAllowedValues> DataValidationMap =
            new Dictionary<DataValidationValues, XlsxAllowedValues>
            {
                { DataValidationValues.None, XlsxAllowedValues.AnyValue },
                { DataValidationValues.Custom, XlsxAllowedValues.Custom },
                { DataValidationValues.Date, XlsxAllowedValues.Date },
                { DataValidationValues.Decimal, XlsxAllowedValues.Decimal },
                { DataValidationValues.List, XlsxAllowedValues.List },
                { DataValidationValues.TextLength, XlsxAllowedValues.TextLength },
                { DataValidationValues.Time, XlsxAllowedValues.Time },
                { DataValidationValues.Whole, XlsxAllowedValues.WholeNumber },
            };

        public static XlsxAllowedValues ToXlsx(this DataValidationValues value)
        {
            return DataValidationMap[value];
        }

        private static readonly IReadOnlyDictionary<DataValidationErrorStyleValues, XlsxErrorStyle> DataValidationErrorStyleMap =
            new Dictionary<DataValidationErrorStyleValues, XlsxErrorStyle>
            {
                { DataValidationErrorStyleValues.Information, XlsxErrorStyle.Information },
                { DataValidationErrorStyleValues.Warning, XlsxErrorStyle.Warning },
                { DataValidationErrorStyleValues.Stop, XlsxErrorStyle.Stop },
            };

        public static XlsxErrorStyle ToXlsx(this DataValidationErrorStyleValues value)
        {
            return DataValidationErrorStyleMap[value];
        }

        private static readonly IReadOnlyDictionary<DataValidationOperatorValues, XlsxOperator> DataValidationOperatorMap =
            new Dictionary<DataValidationOperatorValues, XlsxOperator>
            {
                { DataValidationOperatorValues.Between, XlsxOperator.Between },
                { DataValidationOperatorValues.GreaterThanOrEqual, XlsxOperator.EqualOrGreaterThan },
                { DataValidationOperatorValues.LessThanOrEqual, XlsxOperator.EqualOrLessThan },
                { DataValidationOperatorValues.Equal, XlsxOperator.EqualTo },
                { DataValidationOperatorValues.GreaterThan, XlsxOperator.GreaterThan },
                { DataValidationOperatorValues.LessThan, XlsxOperator.LessThan },
                { DataValidationOperatorValues.NotBetween, XlsxOperator.NotBetween },
                { DataValidationOperatorValues.NotEqual, XlsxOperator.NotEqualTo },
            };

        public static XlsxOperator ToXlsx(this DataValidationOperatorValues value)
        {
            return DataValidationOperatorMap[value];
        }

        private static readonly IReadOnlyDictionary<SheetStateValues, XlsxWorksheetVisibility> SheetStateMap =
            new Dictionary<SheetStateValues, XlsxWorksheetVisibility>
            {
                { SheetStateValues.Visible, XlsxWorksheetVisibility.Visible },
                { SheetStateValues.Hidden, XlsxWorksheetVisibility.Hidden },
                { SheetStateValues.VeryHidden, XlsxWorksheetVisibility.VeryHidden },
            };

        public static XlsxWorksheetVisibility ToXlsx(this SheetStateValues value)
        {
            return SheetStateMap[value];
        }

        private static readonly IReadOnlyDictionary<PhoneticAlignmentValues, XlsxPhoneticAlignment> PhoneticAlignmentMap =
            new Dictionary<PhoneticAlignmentValues, XlsxPhoneticAlignment>
            {
                { PhoneticAlignmentValues.Center, XlsxPhoneticAlignment.Center },
                { PhoneticAlignmentValues.Distributed, XlsxPhoneticAlignment.Distributed },
                { PhoneticAlignmentValues.Left, XlsxPhoneticAlignment.Left },
                { PhoneticAlignmentValues.NoControl, XlsxPhoneticAlignment.NoControl },
            };


        public static XlsxPhoneticAlignment ToXlsx(this PhoneticAlignmentValues value)
        {
            return PhoneticAlignmentMap[value];
        }

        private static readonly IReadOnlyDictionary<PhoneticValues, XlsxPhoneticType> PhoneticMap =
            new Dictionary<PhoneticValues, XlsxPhoneticType>
            {
                { PhoneticValues.FullWidthKatakana, XlsxPhoneticType.FullWidthKatakana },
                { PhoneticValues.HalfWidthKatakana, XlsxPhoneticType.HalfWidthKatakana },
                { PhoneticValues.Hiragana, XlsxPhoneticType.Hiragana },
                { PhoneticValues.NoConversion, XlsxPhoneticType.NoConversion },
            };

        public static XlsxPhoneticType ToXlsx(this PhoneticValues value)
        {
            return PhoneticMap[value];
        }

        private static readonly IReadOnlyDictionary<DataConsolidateFunctionValues, XlsxPivotSummary> DataConsolidateFunctionMap =
            new Dictionary<DataConsolidateFunctionValues, XlsxPivotSummary>
            {
                { DataConsolidateFunctionValues.Sum, XlsxPivotSummary.Sum },
                { DataConsolidateFunctionValues.Count, XlsxPivotSummary.Count },
                { DataConsolidateFunctionValues.Average, XlsxPivotSummary.Average },
                { DataConsolidateFunctionValues.Minimum, XlsxPivotSummary.Minimum },
                { DataConsolidateFunctionValues.Maximum, XlsxPivotSummary.Maximum },
                { DataConsolidateFunctionValues.Product, XlsxPivotSummary.Product },
                { DataConsolidateFunctionValues.CountNumbers, XlsxPivotSummary.CountNumbers },
                { DataConsolidateFunctionValues.StandardDeviation, XlsxPivotSummary.StandardDeviation },
                { DataConsolidateFunctionValues.StandardDeviationP, XlsxPivotSummary.PopulationStandardDeviation },
                { DataConsolidateFunctionValues.Variance, XlsxPivotSummary.Variance },
                { DataConsolidateFunctionValues.VarianceP, XlsxPivotSummary.PopulationVariance },

            };

        public static XlsxPivotSummary ToXlsx(this DataConsolidateFunctionValues value)
        {
            return DataConsolidateFunctionMap[value];
        }

        private static readonly IReadOnlyDictionary<ShowDataAsValues, XlsxPivotCalculation> ShowDataAsMap =
            new Dictionary<ShowDataAsValues, XlsxPivotCalculation>
            {
                { ShowDataAsValues.Normal, XlsxPivotCalculation.Normal },
                { ShowDataAsValues.Difference, XlsxPivotCalculation.DifferenceFrom },
                { ShowDataAsValues.Percent, XlsxPivotCalculation.PercentageOf },
                { ShowDataAsValues.PercentageDifference, XlsxPivotCalculation.PercentageDifferenceFrom },
                { ShowDataAsValues.RunTotal, XlsxPivotCalculation.RunningTotal },
                { ShowDataAsValues.PercentOfRaw, XlsxPivotCalculation.PercentageOfRow }, // There's a typo in the OpenXML SDK =)
                { ShowDataAsValues.PercentOfColumn, XlsxPivotCalculation.PercentageOfColumn },
                { ShowDataAsValues.PercentOfTotal, XlsxPivotCalculation.PercentageOfTotal },
                { ShowDataAsValues.Index, XlsxPivotCalculation.Index },
            };

        public static XlsxPivotCalculation ToXlsx(this ShowDataAsValues value)
        {
            return ShowDataAsMap[value];
        }

        private static readonly IReadOnlyDictionary<SheetViewValues, XlsxSheetViewOptions> SheetViewMap =
            new Dictionary<SheetViewValues, XlsxSheetViewOptions>
            {
                { SheetViewValues.Normal, XlsxSheetViewOptions.Normal },
                { SheetViewValues.PageBreakPreview, XlsxSheetViewOptions.PageBreakPreview },
                { SheetViewValues.PageLayout, XlsxSheetViewOptions.PageLayout },
            };

        public static XlsxSheetViewOptions ToXlsx(this SheetViewValues value)
        {
            return SheetViewMap[value];
        }

        private static readonly IReadOnlyDictionary<ConditionalFormatValues, XlsxConditionalFormatType> ConditionalFormatMap =
            new Dictionary<ConditionalFormatValues, XlsxConditionalFormatType>
            {
                { ConditionalFormatValues.Expression, XlsxConditionalFormatType.Expression },
                { ConditionalFormatValues.CellIs, XlsxConditionalFormatType.CellIs },
                { ConditionalFormatValues.ColorScale, XlsxConditionalFormatType.ColorScale },
                { ConditionalFormatValues.DataBar, XlsxConditionalFormatType.DataBar },
                { ConditionalFormatValues.IconSet, XlsxConditionalFormatType.IconSet },
                { ConditionalFormatValues.Top10, XlsxConditionalFormatType.Top10 },
                { ConditionalFormatValues.UniqueValues, XlsxConditionalFormatType.IsUnique },
                { ConditionalFormatValues.DuplicateValues, XlsxConditionalFormatType.IsDuplicate },
                { ConditionalFormatValues.ContainsText, XlsxConditionalFormatType.ContainsText },
                { ConditionalFormatValues.NotContainsText, XlsxConditionalFormatType.NotContainsText },
                { ConditionalFormatValues.BeginsWith, XlsxConditionalFormatType.StartsWith },
                { ConditionalFormatValues.EndsWith, XlsxConditionalFormatType.EndsWith },
                { ConditionalFormatValues.ContainsBlanks, XlsxConditionalFormatType.IsBlank },
                { ConditionalFormatValues.NotContainsBlanks, XlsxConditionalFormatType.NotBlank },
                { ConditionalFormatValues.ContainsErrors, XlsxConditionalFormatType.IsError },
                { ConditionalFormatValues.NotContainsErrors, XlsxConditionalFormatType.NotError },
                { ConditionalFormatValues.TimePeriod, XlsxConditionalFormatType.TimePeriod },
                { ConditionalFormatValues.AboveAverage, XlsxConditionalFormatType.AboveAverage },
            };

        public static XlsxConditionalFormatType ToXlsx(this ConditionalFormatValues value)
        {
            return ConditionalFormatMap[value];
        }

        private static readonly IReadOnlyDictionary<ConditionalFormatValueObjectValues, XlsxCFContentType> ConditionalFormatValueObjectMap =
            new Dictionary<ConditionalFormatValueObjectValues, XlsxCFContentType>
            {
                { ConditionalFormatValueObjectValues.Number, XlsxCFContentType.Number },
                { ConditionalFormatValueObjectValues.Percent, XlsxCFContentType.Percent },
                { ConditionalFormatValueObjectValues.Max, XlsxCFContentType.Maximum },
                { ConditionalFormatValueObjectValues.Min, XlsxCFContentType.Minimum },
                { ConditionalFormatValueObjectValues.Formula, XlsxCFContentType.Formula },
                { ConditionalFormatValueObjectValues.Percentile, XlsxCFContentType.Percentile },
            };

        public static XlsxCFContentType ToXlsx(this ConditionalFormatValueObjectValues value)
        {
            return ConditionalFormatValueObjectMap[value];
        }

        private static readonly IReadOnlyDictionary<ConditionalFormattingOperatorValues, XlsxCFOperator> ConditionalFormattingOperatorMap =
            new Dictionary<ConditionalFormattingOperatorValues, XlsxCFOperator>
            {
                { ConditionalFormattingOperatorValues.LessThan, XlsxCFOperator.LessThan },
                { ConditionalFormattingOperatorValues.LessThanOrEqual, XlsxCFOperator.EqualOrLessThan },
                { ConditionalFormattingOperatorValues.Equal, XlsxCFOperator.Equal },
                { ConditionalFormattingOperatorValues.NotEqual, XlsxCFOperator.NotEqual },
                { ConditionalFormattingOperatorValues.GreaterThanOrEqual, XlsxCFOperator.EqualOrGreaterThan },
                { ConditionalFormattingOperatorValues.GreaterThan, XlsxCFOperator.GreaterThan },
                { ConditionalFormattingOperatorValues.Between, XlsxCFOperator.Between },
                { ConditionalFormattingOperatorValues.NotBetween, XlsxCFOperator.NotBetween },
                { ConditionalFormattingOperatorValues.ContainsText, XlsxCFOperator.Contains },
                { ConditionalFormattingOperatorValues.NotContains, XlsxCFOperator.NotContains },
                { ConditionalFormattingOperatorValues.BeginsWith, XlsxCFOperator.StartsWith },
                { ConditionalFormattingOperatorValues.EndsWith, XlsxCFOperator.EndsWith },
            };

        public static XlsxCFOperator ToXlsx(this ConditionalFormattingOperatorValues value)
        {
            return ConditionalFormattingOperatorMap[value];
        }

        private static readonly IReadOnlyDictionary<IconSetValues, XlsxIconSetStyle> IconSetMap =
            new Dictionary<IconSetValues, XlsxIconSetStyle>
            {
                { IconSetValues.ThreeArrows, XlsxIconSetStyle.ThreeArrows },
                { IconSetValues.ThreeArrowsGray, XlsxIconSetStyle.ThreeArrowsGray },
                { IconSetValues.ThreeFlags, XlsxIconSetStyle.ThreeFlags },
                { IconSetValues.ThreeTrafficLights1, XlsxIconSetStyle.ThreeTrafficLights1 },
                { IconSetValues.ThreeTrafficLights2, XlsxIconSetStyle.ThreeTrafficLights2 },
                { IconSetValues.ThreeSigns, XlsxIconSetStyle.ThreeSigns },
                { IconSetValues.ThreeSymbols, XlsxIconSetStyle.ThreeSymbols },
                { IconSetValues.ThreeSymbols2, XlsxIconSetStyle.ThreeSymbols2 },
                { IconSetValues.FourArrows, XlsxIconSetStyle.FourArrows },
                { IconSetValues.FourArrowsGray, XlsxIconSetStyle.FourArrowsGray },
                { IconSetValues.FourRedToBlack, XlsxIconSetStyle.FourRedToBlack },
                { IconSetValues.FourRating, XlsxIconSetStyle.FourRating },
                { IconSetValues.FourTrafficLights, XlsxIconSetStyle.FourTrafficLights },
                { IconSetValues.FiveArrows, XlsxIconSetStyle.FiveArrows },
                { IconSetValues.FiveArrowsGray, XlsxIconSetStyle.FiveArrowsGray },
                { IconSetValues.FiveRating, XlsxIconSetStyle.FiveRating },
                { IconSetValues.FiveQuarters, XlsxIconSetStyle.FiveQuarters },
            };

        public static XlsxIconSetStyle ToXlsx(this IconSetValues value)
        {
            return IconSetMap[value];
        }

        private static readonly IReadOnlyDictionary<TimePeriodValues, XlsxTimePeriod> TimePeriodMap =
            new Dictionary<TimePeriodValues, XlsxTimePeriod>
            {
                { TimePeriodValues.Yesterday, XlsxTimePeriod.Yesterday },
                { TimePeriodValues.Today, XlsxTimePeriod.Today },
                { TimePeriodValues.Tomorrow, XlsxTimePeriod.Tomorrow },
                { TimePeriodValues.Last7Days, XlsxTimePeriod.InTheLast7Days },
                { TimePeriodValues.LastWeek, XlsxTimePeriod.LastWeek },
                { TimePeriodValues.ThisWeek, XlsxTimePeriod.ThisWeek },
                { TimePeriodValues.NextWeek, XlsxTimePeriod.NextWeek },
                { TimePeriodValues.LastMonth, XlsxTimePeriod.LastMonth },
                { TimePeriodValues.ThisMonth, XlsxTimePeriod.ThisMonth },
                { TimePeriodValues.NextMonth, XlsxTimePeriod.NextMonth },
            };

        public static XlsxTimePeriod ToXlsx(this TimePeriodValues value)
        {
            return TimePeriodMap[value];
        }

        private static readonly IReadOnlyDictionary<PivotAreaValues, XlsxPivotAreaType> PivotAreaMap =
            new Dictionary<PivotAreaValues, XlsxPivotAreaType>
            {
                { PivotAreaValues.None, XlsxPivotAreaType.None },
                { PivotAreaValues.Normal, XlsxPivotAreaType.Normal },
                { PivotAreaValues.Data, XlsxPivotAreaType.Data },
                { PivotAreaValues.All, XlsxPivotAreaType.All },
                { PivotAreaValues.Origin, XlsxPivotAreaType.Origin },
                { PivotAreaValues.Button, XlsxPivotAreaType.Button },
                { PivotAreaValues.TopRight, XlsxPivotAreaType.TopRight },
                { PivotAreaValues.TopEnd, XlsxPivotAreaType.TopEnd },
            };

        public static XlsxPivotAreaType ToXlsx(this PivotAreaValues value)
        {
            return PivotAreaMap[value];
        }

        private static readonly IReadOnlyDictionary<X14.SparklineTypeValues, XlsxSparklineType> SparklineTypeMap =
            new Dictionary<X14.SparklineTypeValues, XlsxSparklineType>
            {
                { X14.SparklineTypeValues.Line, XlsxSparklineType.Line },
                { X14.SparklineTypeValues.Column, XlsxSparklineType.Column },
                { X14.SparklineTypeValues.Stacked, XlsxSparklineType.Stacked },
            };

        public static XlsxSparklineType ToXlsx(this X14.SparklineTypeValues value)
        {
            return SparklineTypeMap[value];
        }

        private static readonly IReadOnlyDictionary<X14.SparklineAxisMinMaxValues, XlsxSparklineAxisMinMax> SparklineAxisMinMaxMap =
            new Dictionary<X14.SparklineAxisMinMaxValues, XlsxSparklineAxisMinMax>
            {
                { X14.SparklineAxisMinMaxValues.Individual, XlsxSparklineAxisMinMax.Automatic },
                { X14.SparklineAxisMinMaxValues.Group, XlsxSparklineAxisMinMax.SameForAll },
                { X14.SparklineAxisMinMaxValues.Custom, XlsxSparklineAxisMinMax.Custom },
            };

        public static XlsxSparklineAxisMinMax ToXlsx(this X14.SparklineAxisMinMaxValues value)
        {
            return SparklineAxisMinMaxMap[value];
        }

        private static readonly IReadOnlyDictionary<X14.DisplayBlanksAsValues, XlsxDisplayBlanksAsValues> DisplayBlanksAsMap =
            new Dictionary<X14.DisplayBlanksAsValues, XlsxDisplayBlanksAsValues>
            {
                { X14.DisplayBlanksAsValues.Span, XlsxDisplayBlanksAsValues.Interpolate },
                { X14.DisplayBlanksAsValues.Gap, XlsxDisplayBlanksAsValues.NotPlotted },
                { X14.DisplayBlanksAsValues.Zero, XlsxDisplayBlanksAsValues.Zero },
            };

        public static XlsxDisplayBlanksAsValues ToXlsx(this X14.DisplayBlanksAsValues value)
        {
            return DisplayBlanksAsMap[value];
        }

        private static readonly IReadOnlyDictionary<FieldSortValues, XlsxPivotSortType> FieldSortMap =
            new Dictionary<FieldSortValues, XlsxPivotSortType>
            {
                { FieldSortValues.Manual, XlsxPivotSortType.Default },
                { FieldSortValues.Ascending, XlsxPivotSortType.Ascending },
                { FieldSortValues.Descending, XlsxPivotSortType.Descending },
            };

        public static XlsxPivotSortType ToXlsx(this FieldSortValues value)
        {
            return FieldSortMap[value];
        }

        private static readonly IReadOnlyDictionary<PivotTableAxisValues, XlsxPivotAxis> PivotTableAxisMap =
            new Dictionary<PivotTableAxisValues, XlsxPivotAxis>
            {
                { PivotTableAxisValues.AxisRow, XlsxPivotAxis.AxisRow },
                { PivotTableAxisValues.AxisColumn, XlsxPivotAxis.AxisCol },
                { PivotTableAxisValues.AxisPage, XlsxPivotAxis.AxisPage },
                { PivotTableAxisValues.AxisValues, XlsxPivotAxis.AxisValues },
            };

        internal static XlsxPivotAxis ToXlsx(this PivotTableAxisValues value)
        {
            return PivotTableAxisMap[value];
        }

        private static readonly IReadOnlyDictionary<ItemValues, XlsxPivotItemType> ItemMap =
            new Dictionary<ItemValues, XlsxPivotItemType>
            {
                { ItemValues.Data, XlsxPivotItemType.Data },
                { ItemValues.Default, XlsxPivotItemType.Default },
                { ItemValues.Sum, XlsxPivotItemType.Sum },
                { ItemValues.CountA, XlsxPivotItemType.CountA },
                { ItemValues.Average, XlsxPivotItemType.Avg },
                { ItemValues.Maximum, XlsxPivotItemType.Max },
                { ItemValues.Minimum, XlsxPivotItemType.Min },
                { ItemValues.Product, XlsxPivotItemType.Product },
                { ItemValues.Count, XlsxPivotItemType.Count },
                { ItemValues.StandardDeviation, XlsxPivotItemType.StdDev },
                { ItemValues.StandardDeviationP, XlsxPivotItemType.StdDevP },
                { ItemValues.Variance, XlsxPivotItemType.Var },
                { ItemValues.VarianceP, XlsxPivotItemType.VarP },
                { ItemValues.Grand, XlsxPivotItemType.Grand },
                { ItemValues.Blank, XlsxPivotItemType.Blank },
            };

        internal static XlsxPivotItemType ToXlsx(this ItemValues value)
        {
            return ItemMap[value];
        }

        private static readonly IReadOnlyDictionary<FormatActionValues, XlsxPivotFormatAction> FormatActionMap =
            new Dictionary<FormatActionValues, XlsxPivotFormatAction>
            {
                { FormatActionValues.Blank, XlsxPivotFormatAction.Blank },
                { FormatActionValues.Formatting, XlsxPivotFormatAction.Formatting },
            };

        internal static XlsxPivotFormatAction ToXlsx(this FormatActionValues value)
        {
            return FormatActionMap[value];
        }

        private static readonly IReadOnlyDictionary<ScopeValues, XlsxPivotCfScope> ScopeMap =
            new Dictionary<ScopeValues, XlsxPivotCfScope>
            {
                { ScopeValues.Selection, XlsxPivotCfScope.SelectedCells },
                { ScopeValues.Data, XlsxPivotCfScope.DataFields },
                { ScopeValues.Field, XlsxPivotCfScope.FieldIntersections },
            };

        internal static XlsxPivotCfScope ToXlsx(this ScopeValues value)
        {
            return ScopeMap[value];
        }

        private static readonly IReadOnlyDictionary<RuleValues, XlsxPivotCfRuleType> RuleMap =
            new Dictionary<RuleValues, XlsxPivotCfRuleType>
            {
                { RuleValues.None, XlsxPivotCfRuleType.None },
                { RuleValues.All, XlsxPivotCfRuleType.All },
                { RuleValues.Row, XlsxPivotCfRuleType.Row },
                { RuleValues.Column, XlsxPivotCfRuleType.Column },
            };

        internal static XlsxPivotCfRuleType ToXlsx(this RuleValues value)
        {
            return RuleMap[value];
        }

        #endregion To Xlsx
    }
}
