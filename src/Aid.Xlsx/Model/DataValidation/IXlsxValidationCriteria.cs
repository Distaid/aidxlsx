#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.Model.DataValidation;

public interface IXlsxValidationCriteria
{
    void Between(String minValue, String maxValue);

    void Between(IXlsxCell minValue, IXlsxCell maxValue);

    void EqualOrGreaterThan(String value);

    void EqualOrGreaterThan(IXlsxCell cell);

    void EqualOrLessThan(String value);

    void EqualOrLessThan(IXlsxCell cell);

    void EqualTo(String value);

    void EqualTo(IXlsxCell cell);

    void GreaterThan(String value);

    void GreaterThan(IXlsxCell cell);

    void LessThan(String value);

    void LessThan(IXlsxCell cell);

    void NotBetween(String minValue, String maxValue);

    void NotBetween(IXlsxCell minValue, IXlsxCell maxValue);

    void NotEqualTo(String value);

    void NotEqualTo(IXlsxCell cell);
}
