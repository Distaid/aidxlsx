using System;

namespace Aid.Xlsx
{
    internal record XLSortElement(
        Int32 ElementNumber,
        XLSortOrder SortOrder,
        Boolean IgnoreBlanks,
        Boolean MatchCase) : IXLSortElement;
}
