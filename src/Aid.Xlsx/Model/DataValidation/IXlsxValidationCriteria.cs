#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.Model.DataValidation;

public interface IXlsxValidationCriteria
{
    void Between(string minValue, string maxValue);

    void Between(IXlsxCell minValue, IXlsxCell maxValue);

    void EqualOrGreaterThan(string value);

    void EqualOrGreaterThan(IXlsxCell cell);

    void EqualOrLessThan(string value);

    void EqualOrLessThan(IXlsxCell cell);

    void EqualTo(string value);

    void EqualTo(IXlsxCell cell);

    void GreaterThan(string value);

    void GreaterThan(IXlsxCell cell);

    void LessThan(string value);

    void LessThan(IXlsxCell cell);

    void NotBetween(string minValue, string maxValue);

    void NotBetween(IXlsxCell minValue, IXlsxCell maxValue);

    void NotEqualTo(string value);

    void NotEqualTo(IXlsxCell cell);
}
