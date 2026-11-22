#nullable disable

using Aid.Xlsx.Cells;
using Aid.Xlsx.Styles;

namespace Aid.Xlsx.Columns
{
    public interface IXlsxColumns : IEnumerable<IXlsxColumn>
    {
        /// <summary>
        /// Sets the width of all columns.
        /// </summary>
        /// <value>
        /// The width of all columns.
        /// </value>
        Double Width { set; }

        /// <summary>
        /// Deletes all columns and shifts the columns at the right of them accordingly.
        /// </summary>
        void Delete();

        /// <summary>
        /// Adjusts the width of all columns based on its contents.
        /// </summary>
        IXlsxColumns AdjustToContents();

        /// <summary>
        /// Adjusts the width of all columns based on its contents, starting from the startRow.
        /// </summary>
        /// <param name="startRow">The row to start calculating the column width.</param>
        IXlsxColumns AdjustToContents(Int32 startRow);

        /// <summary>
        /// Adjusts the width of all columns based on its contents, starting from the startRow and ending at endRow.
        /// </summary>
        /// <param name="startRow">The row to start calculating the column width.</param>
        /// <param name="endRow">The row to end calculating the column width.</param>
        IXlsxColumns AdjustToContents(Int32 startRow, Int32 endRow);

        IXlsxColumns AdjustToContents(Double minWidth, Double maxWidth);

        IXlsxColumns AdjustToContents(Int32 startRow, Double minWidth, Double maxWidth);

        IXlsxColumns AdjustToContents(Int32 startRow, Int32 endRow, Double minWidth, Double maxWidth);

        /// <summary>
        /// Hides all columns.
        /// </summary>
        void Hide();

        /// <summary>Unhides all columns.</summary>
        void Unhide();

        /// <summary>
        /// Increments the outline level of all columns by 1.
        /// </summary>
        void Group();

        /// <summary>
        /// Increments the outline level of all columns by 1.
        /// </summary>
        /// <param name="collapse">If set to <c>true</c> the columns will be shown collapsed.</param>
        void Group(Boolean collapse);

        /// <summary>
        /// Sets outline level for all columns.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        void Group(Int32 outlineLevel);

        /// <summary>
        /// Sets outline level for all columns.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        /// <param name="collapse">If set to <c>true</c> the columns will be shown collapsed.</param>
        void Group(Int32 outlineLevel, Boolean collapse);

        /// <summary>
        /// Decrements the outline level of all columns by 1.
        /// </summary>
        void Ungroup();

        /// <summary>
        /// Decrements the outline level of all columns by 1.
        /// </summary>
        /// <param name="fromAll">If set to <c>true</c> it will remove the columns from all outline levels.</param>
        void Ungroup(Boolean fromAll);

        /// <summary>
        /// Show all columns as collapsed.
        /// </summary>
        void Collapse();

        /// <summary>Expands all columns (if they're collapsed).</summary>
        void Expand();

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
        /// <param name="includeFormats">if set to <c>true</c> will return all cells with a value or a style different than the default.</param>
        IXlsxCells CellsUsed(Boolean includeFormats);

        IXlsxStyle Style { get; set; }

        /// <summary>
        /// Adds a vertical page break after these columns.
        /// </summary>
        IXlsxColumns AddVerticalPageBreaks();

        /// <summary>
        /// Clears the contents of these columns.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXlsxColumns Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

        void Select();
    }
}
