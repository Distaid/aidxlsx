#nullable disable

namespace Aid.Xlsx.Model.Grid.Cells;

[Flags]
internal enum XlsxCellCopyOptions
{
    None = 0,
    Values = 1 << 1,
    Styles = 1 << 2,
    ConditionalFormats = 1 << 3,
    DataValidations = 1 << 4,
    Sparklines = 1 << 5,
    All = Values | Styles | ConditionalFormats | DataValidations | Sparklines
}
