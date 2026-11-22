#nullable disable

// Keep this file CodeMaid organised and cleaned
using System.Globalization;
using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.Cells;

namespace Aid.Xlsx.AutoFilters
{
    internal enum XlsxConnector { And, Or }

    internal enum XlsxFilterOperator { Equal, NotEqual, GreaterThan, LessThan, EqualOrGreaterThan, EqualOrLessThan }

    /// <summary>
    /// A single filter condition for auto filter.
    /// </summary>
    internal class XlsxFilter
    {
        private XlsxFilter()
        {
        }

        public XlsxConnector Connector { get; set; }

        public XlsxDateTimeGrouping DateTimeGrouping { get; set; }

        /// <summary>
        /// Value for <see cref="XlsxFilterType.Custom"/> that is compared using <see cref="Operator"/>.
        /// </summary>
        public XlsxCellValue CustomValue { get; init; }

        public Func<IXlsxCell, XlsxFilterColumn, bool> Condition { get; init; }

        public XlsxFilterOperator Operator { get; init; } = XlsxFilterOperator.Equal;

        /// <summary>
        /// Value for <see cref="XlsxFilterType.Regular"/> filter.
        /// </summary>
        public Object Value { get; set; }

        internal static XlsxFilter CreateCustomFilter(XlsxCellValue value, XlsxFilterOperator op, XlsxConnector connector)
        {
            // Keep in closure, so it doesn't have to be checked for every cell.
            var comparer = StringComparer.CurrentCultureIgnoreCase;
            return new XlsxFilter
            {
                CustomValue = value,
                Operator = op,
                Connector = connector,
                Condition = (cell, _) => CustomFilterSatisfied(cell.CachedValue, op, value, comparer),
            };
        }

        internal static XlsxFilter CreateCustomPatternFilter(string filterValue, bool match, XlsxConnector connector)
        {
            // Excel really parses value in current culture to detect type (e.g. 1,00 is detected as a number in cs-CZ).
            var testValue = XlsxCellValue.FromText(filterValue, CultureInfo.CurrentCulture);
            if (testValue.Type == XlsxDataType.Text)
            {
                // Custom filter Equal matches strings with a pattern. Custom uses it mostly for filters like begin-with (e.g. `ABC*`).
                var wildcard = filterValue;
                return new XlsxFilter
                {
                    CustomValue = wildcard,
                    Operator = match ? XlsxFilterOperator.Equal : XlsxFilterOperator.NotEqual,
                    Connector = connector,
                    Condition = match ? (c, _) => TextMatchesWildcard(wildcard, c) : (c, _) => !TextMatchesWildcard(wildcard, c),
                };
            }

            // Keep in closure, so it doesn't have to be checked for every cell.
            var comparer = StringComparer.CurrentCultureIgnoreCase;
            return new XlsxFilter
            {
                CustomValue = filterValue,
                Operator = match ? XlsxFilterOperator.Equal : XlsxFilterOperator.NotEqual,
                Connector = connector,
                Condition = match ? (c, _) => ContentMatches(c, filterValue)
                    : (c, _) => !CustomFilterSatisfied(c.CachedValue, XlsxFilterOperator.Equal, testValue, comparer),
            };

            static bool TextMatchesWildcard(string pattern, IXlsxCell cell)
            {
                var cachedValue = cell.CachedValue;
                if (!cachedValue.IsText)
                    return false;

                var wildcard = new Wildcard(pattern);
                var position = wildcard.Search(cachedValue.GetText().AsSpan());
                return position >= 0;
            }
        }

        internal static XlsxFilter CreateRegularFilter(string filterValue)
        {

            return new XlsxFilter
            {
                Value = filterValue,
                Operator = XlsxFilterOperator.Equal,
                Connector = XlsxConnector.Or,
                Condition = (cell, _) => ContentMatches(cell, filterValue)
            };
        }

        internal static XlsxFilter CreateDateGroupFilter(DateTime date, XlsxDateTimeGrouping dateTimeGrouping)
        {
            return new XlsxFilter
            {
                Value = date,
                Condition = HasSameGroup,
                Operator = XlsxFilterOperator.Equal,
                Connector = XlsxConnector.Or,
                DateTimeGrouping = dateTimeGrouping
            };

            bool HasSameGroup(IXlsxCell cell, XlsxFilterColumn _)
            {
                var cachedValue = cell.CachedValue;
                return cachedValue.IsDateTime && IsMatch(date, cachedValue.GetDateTime(), dateTimeGrouping);
            }

            static Boolean IsMatch(DateTime date1, DateTime date2, XlsxDateTimeGrouping dateTimeGrouping)
            {
                Boolean isMatch = true;
                if (dateTimeGrouping >= XlsxDateTimeGrouping.Year) isMatch &= date1.Year.Equals(date2.Year);
                if (isMatch && dateTimeGrouping >= XlsxDateTimeGrouping.Month) isMatch &= date1.Month.Equals(date2.Month);
                if (isMatch && dateTimeGrouping >= XlsxDateTimeGrouping.Day) isMatch &= date1.Day.Equals(date2.Day);
                if (isMatch && dateTimeGrouping >= XlsxDateTimeGrouping.Hour) isMatch &= date1.Hour.Equals(date2.Hour);
                if (isMatch && dateTimeGrouping >= XlsxDateTimeGrouping.Minute) isMatch &= date1.Minute.Equals(date2.Minute);
                if (isMatch && dateTimeGrouping >= XlsxDateTimeGrouping.Second) isMatch &= date1.Second.Equals(date2.Second);

                return isMatch;
            }
        }

        private static bool ContentMatches(IXlsxCell cell, string filterValue)
        {
            // IXlsxCell.GetFormattedString() could trigger formula evaluation.
            var cachedValue = cell.CachedValue;
            var formattedString = ((XlsxCell)cell).GetFormattedString(cachedValue);
            return formattedString.Equals(filterValue, StringComparison.OrdinalIgnoreCase);
        }

        private static bool CustomFilterSatisfied(XlsxCellValue cellValue, XlsxFilterOperator op, XlsxCellValue filterValue, StringComparer textComparer)
        {
            // Blanks are rather strange case. Excel parsing logic for custom filter value into
            // XlsxCellValue is very inconsistent. E.g. 'does not equal' for empty string ignores
            // blanks and empty strings.
            // For custom compare filters, blank never matches.
            if (cellValue.IsBlank || filterValue.IsBlank)
                return false;

            if (cellValue.Type != filterValue.Type)
            {
                // Types are different, but could still be unified numbers and thus comparable.
                if (!(cellValue.IsUnifiedNumber && filterValue.IsUnifiedNumber))
                    return false;
            }

            // Note that custom filter even error values, basically everything as a number.
            var comparison = cellValue.Type switch
            {
                XlsxDataType.Text => textComparer.Compare(cellValue.GetText(), filterValue.GetText()),
                XlsxDataType.Boolean => cellValue.GetBoolean().CompareTo(filterValue.GetBoolean()),
                XlsxDataType.Error => cellValue.GetError().CompareTo(filterValue.GetError()),
                _ => cellValue.GetUnifiedNumber().CompareTo(filterValue.GetUnifiedNumber())
            };

            // !!! Deviation from Excel !!!
            // Excel interprets custom filter with `equal` operator (and *only* equal operator) as
            // comparison of formatted string of a cell value with wildcard represented by custom
            // value filter value.
            // We do the sane thing and compare them for equality, so $10 is equal to 10.
            return op switch
            {
                XlsxFilterOperator.LessThan => comparison < 0,
                XlsxFilterOperator.EqualOrLessThan => comparison <= 0,
                XlsxFilterOperator.Equal => comparison == 0,
                XlsxFilterOperator.NotEqual => comparison != 0,
                XlsxFilterOperator.EqualOrGreaterThan => comparison >= 0,
                XlsxFilterOperator.GreaterThan => comparison > 0,
                _ => throw new NotSupportedException(),
            };
        }

        internal static XlsxFilter CreateTopBottom(bool takeTop, int percentsOrItemCount)
        {
            bool TopFilter(IXlsxCell cell, XlsxFilterColumn filterColumn)
            {
                var cachedValue = cell.CachedValue;
                return cachedValue.IsUnifiedNumber && cachedValue.GetUnifiedNumber() >= filterColumn.TopBottomFilterValue;
            }
            bool BottomFilter(IXlsxCell cell, XlsxFilterColumn filterColumn)
            {
                var cachedValue = cell.CachedValue;
                return cachedValue.IsUnifiedNumber && cachedValue.GetUnifiedNumber() <= filterColumn.TopBottomFilterValue;
            }

            return new XlsxFilter
            {
                Value = percentsOrItemCount,
                Operator = XlsxFilterOperator.Equal,
                Connector = XlsxConnector.Or,
                Condition = takeTop ? TopFilter : BottomFilter,
            };
        }

        internal static XlsxFilter CreateAverage(double initialAverage, bool aboveAverage)
        {
            bool AboveAverage(IXlsxCell cell, XlsxFilterColumn filterColumn)
            {
                var cachedValue = cell.CachedValue;
                var average = filterColumn.DynamicValue;
                return cachedValue.IsUnifiedNumber && cachedValue.GetUnifiedNumber() > average;
            }
            bool BelowAverage(IXlsxCell cell, XlsxFilterColumn filterColumn)
            {
                var cachedValue = cell.CachedValue;
                var average = filterColumn.DynamicValue;
                return cachedValue.IsUnifiedNumber && cachedValue.GetUnifiedNumber() < average;
            }

            return new XlsxFilter
            {
                Value = initialAverage,
                Operator = XlsxFilterOperator.Equal,
                Connector = XlsxConnector.Or,
                Condition = aboveAverage ? AboveAverage : BelowAverage,
            };
        }
    }
}
