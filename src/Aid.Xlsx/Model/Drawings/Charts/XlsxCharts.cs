#nullable disable

namespace Aid.Xlsx.Model.Drawings.Charts;

internal class XlsxCharts : IXlsxCharts
{
    private readonly List<IXlsxChart> charts = [];
    public IEnumerator<IXlsxChart> GetEnumerator()
    {
        return charts.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public void Add(IXlsxChart chart)
    {
        charts.Add(chart);
    }
}
