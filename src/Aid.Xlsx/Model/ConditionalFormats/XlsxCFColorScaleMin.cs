#nullable disable

using Aid.Xlsx.Calculation;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.ConditionalFormats;

internal class XlsxCFColorScaleMin : IXlsxCFColorScaleMin
{
    private readonly XlsxConditionalFormat _conditionalFormat;
    public XlsxCFColorScaleMin(XlsxConditionalFormat conditionalFormat)
    {
        _conditionalFormat = conditionalFormat;
    }

    public IXlsxCFColorScaleMid Minimum(XlsxCFContentType type, String value, XlsxColor color)
    {
        _conditionalFormat.Values.Initialize(new XlsxFormula { Value = value });
        _conditionalFormat.Colors.Initialize(color);
        _conditionalFormat.ContentTypes.Initialize(type);
        return new XlsxCFColorScaleMid(_conditionalFormat);
    }
    public IXlsxCFColorScaleMid Minimum(XlsxCFContentType type, Double value, XlsxColor color)
    {
        return Minimum(type, value.ToInvariantString(), color);
    }

    public IXlsxCFColorScaleMid LowestValue(XlsxColor color)
    {
        _conditionalFormat.Values.Initialize(null);
        _conditionalFormat.Colors.Initialize(color);
        _conditionalFormat.ContentTypes.Initialize(XlsxCFContentType.Minimum);
        return new XlsxCFColorScaleMid(_conditionalFormat);
    }
}
