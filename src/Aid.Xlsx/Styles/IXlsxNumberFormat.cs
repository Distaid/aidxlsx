namespace Aid.Xlsx.Styles;

public interface IXlsxNumberFormat : IXlsxNumberFormatBase, IEquatable<IXlsxNumberFormatBase>
{
    /// <summary>
    /// Sets a predefined number format. Use members of <see cref="XlsxPredefinedFormat.General"/>,
    /// <see cref="XlsxPredefinedFormat.Number"/> or <see cref="XlsxPredefinedFormat.DateTime"/>
    /// casted to <c>int</c>.
    /// <example>
    /// <code>
    ///   cell.SetNumberFormatId((int)XlsxPredefinedFormat.Number.Precision2);
    /// </code>
    /// </example>
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The passed value is not a predefined format.</exception>
    IXlsxStyle SetNumberFormatId(int value);

    /// <summary>
    /// Sets the number format.
    /// </summary>
    IXlsxStyle SetFormat(string value);
}
