namespace Aid.Xlsx.Model.Grid.Ranges.Sort;

internal record XlsxSortElement(
    Int32 ElementNumber,
    XlsxSortOrder SortOrder,
    Boolean IgnoreBlanks,
    Boolean MatchCase) : IXlsxSortElement;
