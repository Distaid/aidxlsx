#nullable disable

using Aid.Xlsx.Formulas.References.Transforms;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Grid.Coordinates;

namespace Aid.Xlsx.Calculation;

public class XlsxFormula
{
    public XlsxFormula()
    { }

    public XlsxFormula(XlsxFormula defaultFormula)
    {
        _value = defaultFormula._value;
        IsFormula = defaultFormula.IsFormula;
    }

    public XlsxFormula(String value)
    {
        Value = value;
    }

    public XlsxFormula(double value)
    {
        Value = value.ToInvariantString();
    }

    public XlsxFormula(int value)
    {
        Value = value.ToInvariantString();
    }

    internal String _value;
    public String Value
    {
        get { return _value; }
        set
        {
            if (value == null)
            {
                _value = String.Empty;
            }
            else
            {
                _value = value.Trim();
                IsFormula = !String.IsNullOrWhiteSpace(_value) && _value.TrimStart()[0] == '=';
                if (IsFormula)
                    _value = _value[1..];
            }


        }
    }

    public Boolean IsFormula { get; internal set; }

    internal XlsxFormula GetAdjustedCopy(Point sourceAnchor, Point targetAnchor)
    {
        if (!IsFormula)
            return new XlsxFormula(this);

        var formulaR1C1 = FormulaConverter.ToR1C1(Value, sourceAnchor.Row, sourceAnchor.Column);
        var formulaA1 = FormulaConverter.ToA1(formulaR1C1, targetAnchor.Row, targetAnchor.Column);
        return new XlsxFormula { _value = formulaA1, IsFormula = true };
    }
}
