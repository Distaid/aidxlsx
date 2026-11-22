#nullable disable

namespace Aid.Xlsx.Ranges
{
    internal enum XlsxRangeType : byte
    {
        Range = 0,
        Column = 1,
        Row = 2,
        RangeColumn = 3,
        RangeRow = 4,
        Table = 5,
        Worksheet = 6
    }
}
