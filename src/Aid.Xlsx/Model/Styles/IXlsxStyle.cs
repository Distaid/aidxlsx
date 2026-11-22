#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Cells;

namespace Aid.Xlsx.Styles
{
    public interface IXlsxStyle : IEquatable<IXlsxStyle>
    {
        IXlsxAlignment Alignment { get; set; }

        IXlsxBorder Border { get; set; }

        IXlsxNumberFormat DateFormat { get; }

        IXlsxFill Fill { get; set; }

        IXlsxFont Font { get; set; }

        /// <summary>
        /// Should the text values of a cell saved to the file be prefixed by a quote (<c>'</c>) character?
        /// Has no effect if cell values is not a <see cref="XlsxDataType.Text"/>. Doesn't affect values during runtime,
        /// text values are returned without quote.
        /// </summary>
        Boolean IncludeQuotePrefix { get; set; }

        IXlsxNumberFormat NumberFormat { get; set; }

        IXlsxProtection Protection { get; set; }

        IXlsxStyle SetIncludeQuotePrefix(Boolean includeQuotePrefix = true);
    }
}
