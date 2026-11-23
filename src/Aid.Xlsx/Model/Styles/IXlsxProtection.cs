using Aid.Xlsx.Model.Protection;

namespace Aid.Xlsx.Model.Styles;

public interface IXlsxProtection : IEquatable<IXlsxProtection>
{
    /// <summary>
    /// Get or set whether a cell is locked.
    /// </summary>
    /// <remarks>
    /// When a cell is locked and its sheet is protected (<c>IXlsxWorksheet.Protection</c>), then
    /// the protected operations (<see cref="XlsxSheetProtectionElements"/>) are prohibited on
    /// the cell.
    /// </remarks>
    bool Locked { get; set; }

    /// <summary>
    /// Get or set whether cell is a hidden.
    /// </summary>
    /// <value>
    /// When a cell is hidden and its sheet is protected (<c>IXlsxWorksheet.Protection</c>), then
    /// application will display content of a cell in a grid, but it will hide formulas bar
    /// content (i.e. formulas or even plain text).
    /// </value>
    bool Hidden { get; set; }

    /// <summary>
    /// Set cell as locked.
    /// </summary>
    /// <inheritdoc cref="Locked"/>
    IXlsxStyle SetLocked();

    /// <summary>
    /// Set whether a cell is locked.
    /// </summary>
    /// <inheritdoc cref="Locked"/>
    IXlsxStyle SetLocked(bool value);

    /// <summary>
    /// Set cell as hidden.
    /// </summary>
    /// <inheritdoc cref="Hidden"/>
    IXlsxStyle SetHidden();

    /// <summary>
    /// Set whether a cell is hidden.
    /// </summary>
    /// <inheritdoc cref="Hidden"/>
    IXlsxStyle SetHidden(bool value);
}
