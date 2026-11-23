#nullable disable

using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;

namespace Aid.Xlsx.Model.PageSetup;

public interface IXlsxPrintAreas : IEnumerable<IXlsxRange>
{
    /// <summary>Removes the print areas from the worksheet.</summary>
    void Clear();

    /// <summary>Adds a range to the print areas.</summary>
    /// <param name="firstCellRow">   The first cell row.</param>
    /// <param name="firstCellColumn">The first cell column.</param>
    /// <param name="lastCellRow">    The last cell row.</param>
    /// <param name="lastCellColumn"> The last cell column.</param>
    void Add(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn);

    /// <summary>Adds a range to the print areas.</summary>
    /// <param name="rangeAddress">The range address to add.</param>
    void Add(string rangeAddress);

    /// <summary>Adds a range to the print areas.</summary>
    /// <param name="firstCellAddress">The first cell address.</param>
    /// <param name="lastCellAddress"> The last cell address.</param>
    void Add(string firstCellAddress, string lastCellAddress);

    /// <summary>Adds a range to the print areas.</summary>
    /// <param name="firstCellAddress">The first cell address.</param>
    /// <param name="lastCellAddress"> The last cell address.</param>
    void Add(IXlsxAddress firstCellAddress, IXlsxAddress lastCellAddress);
}
