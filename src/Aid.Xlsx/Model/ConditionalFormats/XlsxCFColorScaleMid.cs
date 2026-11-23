#nullable disable

using Aid.Xlsx.Calculation;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.ConditionalFormats;

internal class XlsxCFColorScaleMid(XlsxConditionalFormat conditionalFormat) : IXlsxCFColorScaleMid
{
    private readonly XlsxConditionalFormat _conditionalFormat = conditionalFormat;

    public IXlsxCFColorScaleMax Midpoint(XlsxCFContentType type, string value, XlsxColor color)
    {
        _conditionalFormat.Values.Add(new XlsxFormula { Value = value });
        _conditionalFormat.Colors.Add(color);
        _conditionalFormat.ContentTypes.Add(type);
        return new XlsxCFColorScaleMax(_conditionalFormat);
    }
    public IXlsxCFColorScaleMax Midpoint(XlsxCFContentType type, double value, XlsxColor color)
    {
        return Midpoint(type, value.ToInvariantString(), color);
    }
    public void Maximum(XlsxCFContentType type, string value, XlsxColor color)
    {
        Midpoint(type, value, color);
    }
    public void Maximum(XlsxCFContentType type, double value, XlsxColor color)
    {
        Maximum(type, value.ToInvariantString(), color);
    }
    public void HighestValue(XlsxColor color)
    {
        _conditionalFormat.Values.Initialize(null);
        _conditionalFormat.Colors.Add(color);
        _conditionalFormat.ContentTypes.Add(XlsxCFContentType.Maximum);
    }
}
