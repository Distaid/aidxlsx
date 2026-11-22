#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Styles;

namespace Aid.Xlsx.Model.Tables;

public interface IXlsxTableRows : IEnumerable<IXlsxTableRow>
{
    IXlsxStyle Style { get; set; }

    /// <summary>
    /// Adds a table row to this group.
    /// </summary>
    /// <param name="tableRow">The row table to add.</param>
    void Add(IXlsxTableRow tableRow);

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
    /// Clears the contents of these rows.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    IXlsxTableRows Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All);

    void Delete();

    void Select();
}
