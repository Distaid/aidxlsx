#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Infrastructure.Utilities;

namespace Aid.Xlsx.Infrastructure.Extensions;

internal static class DoubleExtensions
{
    extension(double value)
    {
        public double SaveRound()
        {
            return Math.Round(value, 6);
        }

        public TimeSpan ToSerialTimeSpan()
        {
            return XlsxHelper.GetTimeSpan(value);
        }

        public DateTime ToSerialDateTime()
        {
            if (value >= 61.0)
                return DateTime.FromOADate(value);
            if (value <= 60.0)
                return DateTime.FromOADate(value + 1);

            throw new ArgumentException($"Serial date 60 is on a leap year of 1900 - date that doesn't exist and isn't representable in DateTime.");
        }

        /// <summary>
        /// Round the number to the integer.
        /// </summary>
        /// <remarks>A helper method to avoid need to specify the midpoint rounding and casting each time.</remarks>
        public int RoundToInt()
        {
            return (int)Math.Round(value, MidpointRounding.AwayFromZero);
        }

        /// <summary>
        /// Round the number to specified number of digits.
        /// </summary>
        /// <remarks>A helper method to avoid need to specify the midpoint rounding each time.</remarks>
        public double Round(int digits)
        {
            return Math.Round(value, digits, MidpointRounding.AwayFromZero);
        }
    }
}
