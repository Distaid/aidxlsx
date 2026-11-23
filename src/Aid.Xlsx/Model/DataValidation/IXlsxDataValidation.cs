#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Model.Grid.Ranges;

namespace Aid.Xlsx.Model.DataValidation;

public enum XlsxAllowedValues { AnyValue, WholeNumber, Decimal, Date, Time, TextLength, List, Custom }

public enum XlsxErrorStyle { Stop, Warning, Information }

public enum XlsxOperator { EqualTo, NotEqualTo, GreaterThan, LessThan, EqualOrGreaterThan, EqualOrLessThan, Between, NotBetween }

public interface IXlsxDataValidation
{
    XlsxAllowedValues AllowedValues { get; set; }

    XlsxDateCriteria Date { get; }

    XlsxDecimalCriteria Decimal { get; }

    string ErrorMessage { get; set; }

    XlsxErrorStyle ErrorStyle { get; set; }

    string ErrorTitle { get; set; }

    bool IgnoreBlanks { get; set; }

    bool InCellDropdown { get; set; }

    string InputMessage { get; set; }

    string InputTitle { get; set; }

    string MaxValue { get; set; }

    string MinValue { get; set; }

    XlsxOperator Operator { get; set; }

    /// <summary>
    /// A collection of ranges the data validation rule applies too.
    /// </summary>
    IEnumerable<IXlsxRange> Ranges { get; }

    bool ShowErrorMessage { get; set; }

    //void Delete();
    //void CopyFrom(IXlsxDataValidation dataValidation);
    bool ShowInputMessage { get; set; }

    XlsxTextLengthCriteria TextLength { get; }

    XlsxTimeCriteria Time { get; }

    string Value { get; set; }

    XlsxWholeNumberCriteria WholeNumber { get; }

    /// <summary>
    /// Add a range to the collection of ranges this rule applies to.
    /// If the specified range does not belong to the worksheet of the data validation
    /// rule it is transferred to the target worksheet.
    /// </summary>
    /// <param name="range">A range to add.</param>
    void AddRange(IXlsxRange range);

    /// <summary>
    /// Add a collection of ranges to the collection of ranges this rule applies to.
    /// Ranges that do not belong to the worksheet of the data validation
    /// rule are transferred to the target worksheet.
    /// </summary>
    /// <param name="ranges">Ranges to add.</param>
    void AddRanges(IEnumerable<IXlsxRange> ranges);

    void Clear();

    /// <summary>
    /// Detach data validation rule of all ranges it applies to.
    /// </summary>
    void ClearRanges();

    void Custom(string customValidation);

    bool IsDirty();

    void List(string list);

    void List(string list, bool inCellDropdown);

    void List(IXlsxRange range);

    void List(IXlsxRange range, bool inCellDropdown);

    /// <summary>
    /// Remove the specified range from the collection of range this rule applies to.
    /// </summary>
    /// <param name="range">A range to remove.</param>
    bool RemoveRange(IXlsxRange range);
}
