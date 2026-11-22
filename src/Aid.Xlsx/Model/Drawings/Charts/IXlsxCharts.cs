#nullable disable

namespace Aid.Xlsx.Model.Drawings.Charts;

public interface IXlsxCharts : IEnumerable<IXlsxChart>
{
    void Add(IXlsxChart chart);
}
