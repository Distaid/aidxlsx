#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Infrastructure.Extensions;

namespace Aid.Xlsx.Model.DataValidation;

public class XlsxDecimalCriteria(IXlsxDataValidation dataValidation) : XlsxValidationCriteria(dataValidation)
{
    public void Between(double minValue, double maxValue) => base.Between(minValue.ToInvariantString(), maxValue.ToInvariantString());

    public void EqualOrGreaterThan(double value) => base.EqualOrGreaterThan(value.ToInvariantString());

    public void EqualOrLessThan(double value) => base.EqualOrLessThan(value.ToInvariantString());

    public void EqualTo(double value) => base.EqualTo(value.ToInvariantString());

    public void GreaterThan(double value) => base.GreaterThan(value.ToInvariantString());

    public void LessThan(double value) => base.LessThan(value.ToInvariantString());

    public void NotBetween(double minValue, double maxValue) => base.NotBetween(minValue.ToInvariantString(), maxValue.ToInvariantString());

    public void NotEqualTo(double value) => base.NotEqualTo(value.ToInvariantString());
}
