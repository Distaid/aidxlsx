namespace Aid.Xlsx.Model.Grid.Ranges.Sort;

internal record XlsxSortElement(
    int ElementNumber,
    XlsxSortOrder SortOrder,
    bool IgnoreBlanks,
    bool MatchCase) : IXlsxSortElement;
