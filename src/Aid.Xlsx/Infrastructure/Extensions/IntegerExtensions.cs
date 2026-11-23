#nullable disable

// Keep this file CodeMaid organised and cleaned

using System.Diagnostics;
using System.Numerics;

namespace Aid.Xlsx.Infrastructure.Extensions;

internal static class IntegerExtensions
{
    public static bool Between(this int val, int from, int to)
    {
        return val >= from && val <= to;
    }

    /// <summary>
    /// Get index of highest set bit &lt;= to <paramref name="maximalIndex"/> or -1 if no such bit.
    /// </summary>
    internal static int GetHighestSetBitBelow(this uint value, int maximalIndex)
    {
        Debug.Assert(maximalIndex >= 0 && maximalIndex < 32);
        value <<= 31 - maximalIndex;
        return value == 0 ? -1 : maximalIndex - BitOperations.LeadingZeroCount(value);
    }

    /// <summary>
    /// Get index of lowest set bit &gt;= to <paramref name="minimalIndex"/> or -1 if no such bit.
    /// </summary>
    internal static int GetLowestSetBitAbove(this uint value, int minimalIndex)
    {
        value >>= minimalIndex;
        return value == 0 ? -1 : minimalIndex + BitOperations.TrailingZeroCount(value);
    }

    /// <summary>
    /// Get highest set bit index or -1 if no bit is set.
    /// </summary>
    internal static int GetHighestSetBit(this uint value)
    {
        return 31 - BitOperations.LeadingZeroCount(value);
    }
}
