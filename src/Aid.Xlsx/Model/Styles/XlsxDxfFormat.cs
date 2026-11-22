using Aid.Xlsx.Styles.Formatting;

namespace Aid.Xlsx.Styles;

internal partial class XlsxDxFormat
{
    private readonly XlsxWorkbookStyles _styles;
    private readonly IXlsxDxfContainer _container;

    internal XlsxDxFormat(XlsxWorkbookStyles styles, IXlsxDxfContainer container)
    {
        _styles = styles;
        _container = container;
    }

    private XlsxDxfValue Dxf => _container.FormatValue ?? XlsxDxfValue.Empty;

    private XlsxDxfNumberFormat NumberFormat => new(this);

    private XlsxDxfAlignmentFormat Alignment => new(this);

    private XlsxDxfFontFormat Font => new(this);

    private XlsxDxfFillFormat Fill => new(this);

    private XlsxDxfBorderFormat Border => new(this);

    private XlsxDxfProtectionFormat Protection => new(this);

    internal TProperty? Resolve<TComponent, TProperty>(Func<XlsxDxfValue, TComponent> getComponent, Func<TComponent, TProperty?> getProperty)
        where TProperty : struct
    {
        var component = getComponent(Dxf);
        return getProperty(component);
    }

    internal TProperty? Resolve<TComponent, TProperty>(Func<XlsxDxfValue, TComponent> getComponent, Func<TComponent, TProperty?> getProperty)
        where TProperty : class
    {
        var component = getComponent(Dxf);
        return getProperty(component);
    }

    internal void ModifyNumberFormat(XlsxNumberFormat numberFormat)
    {
        _container.FormatValue = _styles.RegisterDxFormat(Dxf with { NumberFormat = numberFormat }); ;
    }

    internal void ModifyFont<T>(Func<XlsxDifferentialFontValue, T, XlsxDifferentialFontValue> modify, T value)
    {
        var modifiedDxf = _styles.GetRegisteredDxFormat(Dxf, dxf =>
        {
            var modifiedFont = modify(dxf.Font, value);
            var modifiedDxf = dxf with { Font = modifiedFont };
            return modifiedDxf;
        });
        _container.FormatValue = modifiedDxf;
    }

    internal void ModifyFill<T>(Func<XlsxDifferentialFillValue, T, XlsxDifferentialFillValue> modify, T value)
    {
        var modifiedDxf = _styles.GetRegisteredDxFormat(Dxf, dxf =>
        {
            var modifiedFill = modify(dxf.Fill, value);
            var modifiedDxf = dxf with { Fill = modifiedFill };
            return modifiedDxf;
        });
        _container.FormatValue = modifiedDxf;
    }

    internal void ModifyAlignment<T>(Func<XlsxDifferentialAlignmentValue, T, XlsxDifferentialAlignmentValue> modify, T value)
    {
        var modifiedDxf = _styles.GetRegisteredDxFormat(Dxf, dxf =>
        {
            var modifiedAlignment = modify(dxf.Alignment, value);
            var modifiedDxf = dxf with { Alignment = modifiedAlignment };
            return modifiedDxf;
        });
        _container.FormatValue = modifiedDxf;
    }

    internal void ModifyBorder<T>(Func<XlsxDifferentialBorderValue, T, XlsxDifferentialBorderValue> modify, T value)
    {
        var modifiedDxf = _styles.GetRegisteredDxFormat(Dxf, dxf =>
        {
            var modifiedBorder = modify(dxf.Border, value);
            var modifiedDxf = dxf with { Border = modifiedBorder };
            return modifiedDxf;
        });
        _container.FormatValue = modifiedDxf;
    }

    internal void ModifyProtection<T>(Func<XlsxDifferentialProtectionValue, T, XlsxDifferentialProtectionValue> modify, T value)
    {
        var modifiedDxf = _styles.GetRegisteredDxFormat(Dxf, dxf =>
        {
            var modifiedProtection = modify(dxf.Protection, value);
            var modifiedDxf = dxf with { Protection = modifiedProtection };
            return modifiedDxf;
        });
        _container.FormatValue = modifiedDxf;
    }

    /// <summary>
    /// A helper method that is used when a style if copied from one object to another.
    /// For example, <c>conditionaFormat.Style = someOtherApi.Style</c>.
    /// </summary>
    internal void SetStyle(IXlsxStyle other)
    {
        if (other is not XlsxDxFormat otherFormat)
            throw new NotSupportedException("Can only copy from dxf.");

        _container.FormatValue = _styles.RegisterDxFormat(otherFormat.Dxf);
    }
}
