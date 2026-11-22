namespace Aid.Xlsx.Ranges.Sort
{
    internal record XlsxSortElement(
        Int32 ElementNumber,
        XlsxSortOrder SortOrder,
        Boolean IgnoreBlanks,
        Boolean MatchCase) : IXlsxSortElement;
}
