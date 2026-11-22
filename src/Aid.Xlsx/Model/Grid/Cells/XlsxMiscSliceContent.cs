using Aid.Xlsx.Model.Comments;

namespace Aid.Xlsx.Model.Grid.Cells;

internal struct XlsxMiscSliceContent
{
    internal XlsxComment? Comment { get; set; }

    internal uint? CellMetaIndex { get; set; }

    internal uint? ValueMetaIndex { get; set; }

    internal bool HasPhonetic { get; set; }
}
