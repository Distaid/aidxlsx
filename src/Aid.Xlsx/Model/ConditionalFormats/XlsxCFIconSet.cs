#nullable disable

using Aid.Xlsx.Calculation;
using Aid.Xlsx.Infrastructure.Extensions;

namespace Aid.Xlsx.Model.ConditionalFormats;

internal class XlsxCFIconSet : IXlsxCFIconSet
{
    private readonly XlsxConditionalFormat _conditionalFormat;
    public XlsxCFIconSet(XlsxConditionalFormat conditionalFormat)
    {
        _conditionalFormat = conditionalFormat;
    }
    public IXlsxCFIconSet AddValue(XlsxCFIconSetOperator setOperator, String value, XlsxCFContentType type)
    {
        _conditionalFormat.IconSetOperators.Add(setOperator);
        _conditionalFormat.Values.Add(new XlsxFormula { Value = value });
        _conditionalFormat.ContentTypes.Add(type);
        return new XlsxCFIconSet(_conditionalFormat);
    }
    public IXlsxCFIconSet AddValue(XlsxCFIconSetOperator setOperator, Double value, XlsxCFContentType type)
    {
        return AddValue(setOperator, value.ToInvariantString(), type);
    }
}
