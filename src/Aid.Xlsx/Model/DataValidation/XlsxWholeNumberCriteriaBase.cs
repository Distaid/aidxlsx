#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Infrastructure.Extensions;

namespace Aid.Xlsx.Model.DataValidation;

public abstract class XlsxWholeNumberCriteriaBase(IXlsxDataValidation dataValidation) : XlsxValidationCriteria(dataValidation)
{
    public void Between(int minValue, int maxValue) => base.Between(minValue.ToInvariantString(), maxValue.ToInvariantString());

    public void EqualOrGreaterThan(int value) => base.EqualOrGreaterThan(value.ToInvariantString());

    public void EqualOrLessThan(int value) => base.EqualOrLessThan(value.ToInvariantString());

    public void EqualTo(int value) => base.EqualTo(value.ToInvariantString());

    public void GreaterThan(int value) => base.GreaterThan(value.ToInvariantString());

    public void LessThan(int value) => base.LessThan(value.ToInvariantString());

    public void NotBetween(int minValue, int maxValue) => base.NotBetween(minValue.ToInvariantString(), maxValue.ToInvariantString());

    public void NotEqualTo(int value) => base.NotEqualTo(value.ToInvariantString());
}
