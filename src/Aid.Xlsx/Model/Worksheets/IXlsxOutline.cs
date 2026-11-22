#nullable disable


namespace Aid.Xlsx.Worksheets
{
    public enum XlsxOutlineSummaryVLocation { Top, Bottom };
    public enum XlsxOutlineSummaryHLocation { Left, Right };
    public interface IXlsxOutline
    {
        XlsxOutlineSummaryVLocation SummaryVLocation { get; set; }
        XlsxOutlineSummaryHLocation SummaryHLocation { get; set; }
    }
}
