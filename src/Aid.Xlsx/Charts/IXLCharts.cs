#nullable disable

using System.Collections.Generic;

namespace Aid.Xlsx
{
    public interface IXLCharts : IEnumerable<IXLChart>
    {
        void Add(IXLChart chart);
    }
}
