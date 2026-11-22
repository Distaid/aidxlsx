#nullable disable


namespace Aid.Xlsx.Model.Worksheets;

internal class XlsxOutline : IXlsxOutline
{
    public XlsxOutline(IXlsxOutline outline)
    {
        if (outline != null)
        {
            SummaryHLocation = outline.SummaryHLocation;
            SummaryVLocation = outline.SummaryVLocation;
        }
    }
    public XlsxOutlineSummaryVLocation SummaryVLocation { get; set; }
    public XlsxOutlineSummaryHLocation SummaryHLocation { get; set; }
}
