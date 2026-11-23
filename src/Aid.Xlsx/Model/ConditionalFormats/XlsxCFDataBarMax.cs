#nullable disable

using Aid.Xlsx.Calculation;
using Aid.Xlsx.Infrastructure.Extensions;

namespace Aid.Xlsx.Model.ConditionalFormats;

internal class XlsxCFDataBarMax(XlsxConditionalFormat conditionalFormat) : IXlsxCFDataBarMax
{
    private readonly XlsxConditionalFormat _conditionalFormat = conditionalFormat;

    public void Maximum(XlsxCFContentType type, String value)
    {
        _conditionalFormat.ContentTypes.Add(type);
        _conditionalFormat.Values.Add(new XlsxFormula { Value = value });
    }
    public void Maximum(XlsxCFContentType type, Double value)
    {
        Maximum(type, value.ToInvariantString());
    }

    public void HighestValue()
    {
        Maximum(XlsxCFContentType.Maximum, "0");
    }
}
