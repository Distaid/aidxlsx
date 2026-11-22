using Aid.Xlsx.Model.ConditionalFormats;

namespace Aid.Xlsx.Model.PivotTables;

/// <summary>
/// Specifies how to apply <see cref="XlsxConditionalFormatType.Top10"/> conditional formatting rule
/// on a pivot table <see cref="XlsxPivotConditionalFormat"/>. Avoid if possible, doesn't seem to
/// work and row/column causes Excel to repair file.
/// </summary>
/// <remarks>18.18.84 ST_Type.</remarks>
internal enum XlsxPivotCfRuleType
{
    All,
    Column,
    None,
    Row
}
