#nullable disable

namespace Aid.Xlsx.Charts
{
    public interface IXlsxCharts : IEnumerable<IXlsxChart>
    {
        void Add(IXlsxChart chart);
    }
}
