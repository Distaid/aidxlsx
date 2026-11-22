#nullable disable

namespace Aid.Xlsx.Cells
{
    [Flags]
    public enum XlsxClearOptions
    {
        Contents = 1 << 0,
        NormalFormats = 1 << 2,
        ConditionalFormats = 1 << 3,
        Comments = 1 << 4,
        DataValidation = 1 << 5,
        MergedRanges = 1 << 6,
        Sparklines = 1 << 7,

        AllFormats = NormalFormats | ConditionalFormats,
        AllContents = Contents | Comments,
        All = Contents | NormalFormats | ConditionalFormats | Comments | DataValidation | MergedRanges | Sparklines
    }

    internal static class XlsxClearOptionsExtensions
    {
        public static XlsxCellsUsedOptions ToCellsUsedOptions(this XlsxClearOptions options)
        {
            return (XlsxCellsUsedOptions)options;
        }
    }
}
