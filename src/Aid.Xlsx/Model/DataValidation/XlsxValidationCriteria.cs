#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.Model.DataValidation;

public abstract class XlsxValidationCriteria(IXlsxDataValidation dataValidation) : IXlsxValidationCriteria
{
    protected IXlsxDataValidation dataValidation = dataValidation;

    #region IXlsxValidationCriteria Members

    public void Between(String minValue, String maxValue)
    {
        dataValidation.MinValue = minValue;
        dataValidation.MaxValue = maxValue;
        dataValidation.Operator = XlsxOperator.Between;
    }

    public void Between(IXlsxCell minValue, IXlsxCell maxValue)
    {
        dataValidation.MinValue = minValue.Address.ToStringFixed();
        dataValidation.MaxValue = maxValue.Address.ToStringFixed();
        dataValidation.Operator = XlsxOperator.Between;
    }

    public void EqualOrGreaterThan(String value)
    {
        dataValidation.Value = value;
        dataValidation.Operator = XlsxOperator.EqualOrGreaterThan;
    }

    public void EqualOrGreaterThan(IXlsxCell cell)
    {
        dataValidation.Value = cell.Address.ToStringFixed();
        dataValidation.Operator = XlsxOperator.EqualOrGreaterThan;
    }

    public void EqualOrLessThan(String value)
    {
        dataValidation.Value = value;
        dataValidation.Operator = XlsxOperator.EqualOrLessThan;
    }

    public void EqualOrLessThan(IXlsxCell cell)
    {
        dataValidation.Value = cell.Address.ToStringFixed();
        dataValidation.Operator = XlsxOperator.EqualOrLessThan;
    }

    public void EqualTo(String value)
    {
        dataValidation.Value = value;
        dataValidation.Operator = XlsxOperator.EqualTo;
    }

    public void EqualTo(IXlsxCell cell)
    {
        dataValidation.Value = cell.Address.ToStringFixed();
        dataValidation.Operator = XlsxOperator.EqualTo;
    }

    public void GreaterThan(String value)
    {
        dataValidation.Value = value;
        dataValidation.Operator = XlsxOperator.GreaterThan;
    }

    public void GreaterThan(IXlsxCell cell)
    {
        dataValidation.Value = cell.Address.ToStringFixed();
        dataValidation.Operator = XlsxOperator.GreaterThan;
    }

    public void LessThan(String value)
    {
        dataValidation.Value = value;
        dataValidation.Operator = XlsxOperator.LessThan;
    }

    public void LessThan(IXlsxCell cell)
    {
        dataValidation.Value = cell.Address.ToStringFixed();
        dataValidation.Operator = XlsxOperator.LessThan;
    }

    public void NotBetween(String minValue, String maxValue)
    {
        dataValidation.MinValue = minValue;
        dataValidation.MaxValue = maxValue;
        dataValidation.Operator = XlsxOperator.NotBetween;
    }

    public void NotBetween(IXlsxCell minValue, IXlsxCell maxValue)
    {
        dataValidation.MinValue = minValue.Address.ToStringFixed();
        dataValidation.MaxValue = maxValue.Address.ToStringFixed();
        dataValidation.Operator = XlsxOperator.NotBetween;
    }

    public void NotEqualTo(String value)
    {
        dataValidation.Value = value;
        dataValidation.Operator = XlsxOperator.NotEqualTo;
    }

    public void NotEqualTo(IXlsxCell cell)
    {
        dataValidation.Value = cell.Address.ToStringFixed();
        dataValidation.Operator = XlsxOperator.NotEqualTo;
    }

    #endregion IXlsxValidationCriteria Members
}
