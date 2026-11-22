using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.DataValidation;

public interface IXlsxDataValidations : IEnumerable<IXlsxDataValidation>
{
    IXlsxWorksheet Worksheet { get; }

    /// <summary>
    /// Add data validation rule to the collection. If the specified rule refers to another
    /// worksheet than the collection, the copy will be created and its ranges will refer
    /// to the worksheet of the collection. Otherwise the original instance will be placed
    /// in the collection.
    /// </summary>
    /// <param name="dataValidation">A data validation rule to add.</param>
    /// <returns>The instance that has actually been added in the collection
    /// (may be a copy of the specified one).</returns>
    IXlsxDataValidation Add(IXlsxDataValidation dataValidation);

    Boolean ContainsSingle(IXlsxRange range);

    void Delete(Predicate<IXlsxDataValidation> predicate);

    /// <summary>
    /// Get all data validation rules applied to ranges that intersect the specified range.
    /// </summary>
    IEnumerable<IXlsxDataValidation> GetAllInRange(IXlsxRangeAddress rangeAddress);

    /// <summary>
    /// Get the data validation rule for the range with the specified address if it exists.
    /// </summary>
    /// <param name="rangeAddress">A range address.</param>
    /// <param name="dataValidation">Data validation rule which ranges collection includes the specified
    /// address. The specified range should be fully covered with the data validation rule.
    /// For example, if the rule is applied to ranges A1:A3,C1:C3 then this method will
    /// return True for ranges A1:A3, C1:C2, A2:A3, and False for ranges A1:C3, A1:C1, etc.</param>
    /// <returns>True is the data validation rule was found, false otherwise.</returns>
    bool TryGet(IXlsxRangeAddress rangeAddress, [NotNullWhen(true)] out IXlsxDataValidation? dataValidation);
}
