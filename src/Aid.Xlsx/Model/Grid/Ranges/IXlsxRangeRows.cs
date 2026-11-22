#nullable disable

using Aid.Xlsx.Cells;
using Aid.Xlsx.Styles;

namespace Aid.Xlsx.Ranges
{
    public interface IXlsxRangeRows : IEnumerable<IXlsxRangeRow>
    {
        /// <summary>
        /// Adds a row range to this group.
        /// </summary>
        /// <param name="rowRange">The row range to add.</param>
        void Add(IXlsxRangeRow rowRange);

        /// <summary>
        /// Returns the collection of cells.
        /// </summary>
        IXlsxCells Cells();

        /// <summary>
        /// Returns the collection of cells that have a value.
        /// </summary>
        IXlsxCells CellsUsed();


        /// <summary>
        /// Returns the collection of cells that have a value.
        /// </summary>
        /// <param name="options">The options to determine whether a cell is used.</param>
        IXlsxCells CellsUsed(XlsxCellsUsedOptions options);

        /// <summary>
        /// Deletes all rows and shifts the rows below them accordingly.
        /// </summary>
        void Delete();

        IXlsxStyle Style { get; set; }

        /// <summary>
        /// Clears the contents of these rows.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXlsxRangeRows Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

        void Select();
    }
}
