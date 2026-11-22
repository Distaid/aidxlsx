using Aid.Xlsx.Tables;

namespace Aid.Xlsx.Styles.Formatting;

/// <summary>
/// A table region that can have applied a differential formatting through
/// a <see cref="XlsxTableTheme"/>. Enum values are ordered based on application priority, from
/// the first applied to the last applied.
/// </summary>
internal enum XlsxTableStyleRegionValues
{
    WholeTable,
    FirstColumnStripe,
    SecondColumnStripe,
    FirstRowStripe,
    SecondRowStripe,
    LastColumn,
    FirstColumn,
    HeaderRow,
    TotalRow,
    FirstHeaderCell,
    LastHeaderCell,
    FirstTotalCell,
    LastTotalCell,
}
