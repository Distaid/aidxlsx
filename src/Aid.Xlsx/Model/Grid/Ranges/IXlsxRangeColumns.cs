#nullable disable

using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Styles;

namespace Aid.Xlsx.Model.Grid.Ranges;

public interface IXlsxRangeColumns : IEnumerable<IXlsxRangeColumn>
{
    /// <summary>
    /// Adds a column range to this group.
    /// </summary>
    /// <param name="columRange">The column range to add.</param>
    void Add(IXlsxRangeColumn columRange);

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
    /// Deletes all columns and shifts the columns at the right of them accordingly.
    /// </summary>
    void Delete();

    IXlsxStyle Style { get; set; }

    /// <summary>
    /// Clears the contents of these columns.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    IXlsxRangeColumns Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

    void Select();
}
