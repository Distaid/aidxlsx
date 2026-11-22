using Aid.Xlsx.Comments;

namespace Aid.Xlsx.Cells
{
    internal struct XlsxMiscSliceContent
    {
        internal XlsxComment? Comment { get; set; }

        internal uint? CellMetaIndex { get; set; }

        internal uint? ValueMetaIndex { get; set; }

        internal bool HasPhonetic { get; set; }
    }
}
