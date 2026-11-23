#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Infrastructure.Extensions;

namespace Aid.Xlsx.Model.DataValidation;

public class XlsxDateCriteria(IXlsxDataValidation dataValidation) : XlsxValidationCriteria(dataValidation)
{
    public void Between(DateTime minValue, DateTime maxValue)
    {
        dataValidation.MinValue = minValue.ToOADate().ToInvariantString();
        dataValidation.MaxValue = maxValue.ToOADate().ToInvariantString();
        dataValidation.Operator = XlsxOperator.Between;
    }

    public void EqualOrGreaterThan(DateTime value)
    {
        dataValidation.Value = value.ToOADate().ToInvariantString();
        dataValidation.Operator = XlsxOperator.EqualOrGreaterThan;
    }

    public void EqualOrLessThan(DateTime value)
    {
        dataValidation.Value = value.ToOADate().ToInvariantString();
        dataValidation.Operator = XlsxOperator.EqualOrLessThan;
    }

    public void EqualTo(DateTime value)
    {
        dataValidation.Value = value.ToOADate().ToInvariantString();
        dataValidation.Operator = XlsxOperator.EqualTo;
    }

    public void GreaterThan(DateTime value)
    {
        dataValidation.Value = value.ToOADate().ToInvariantString();
        dataValidation.Operator = XlsxOperator.GreaterThan;
    }

    public void LessThan(DateTime value)
    {
        dataValidation.Value = value.ToOADate().ToInvariantString();
        dataValidation.Operator = XlsxOperator.LessThan;
    }

    public void NotBetween(DateTime minValue, DateTime maxValue)
    {
        dataValidation.MinValue = minValue.ToOADate().ToInvariantString();
        dataValidation.MaxValue = maxValue.ToOADate().ToInvariantString();
        dataValidation.Operator = XlsxOperator.NotBetween;
    }

    public void NotEqualTo(DateTime value)
    {
        dataValidation.Value = value.ToOADate().ToInvariantString();
        dataValidation.Operator = XlsxOperator.NotEqualTo;
    }
}
