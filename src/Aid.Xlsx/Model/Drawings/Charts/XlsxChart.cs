#nullable disable

using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Drawings.Charts;

internal enum XlsxChartTypeCategory { Bar3D }
internal enum XlsxBarOrientation { Vertical, Horizontal }
internal enum XlsxBarGrouping { Clustered, Percent, Stacked, Standard }
internal class XlsxChart : XlsxDrawing<IXlsxChart>, IXlsxChart
{
    internal IXlsxWorksheet worksheet;
    public XlsxChart(XlsxWorksheet worksheet)
    {
        Container = this;
        this.worksheet = worksheet;
        int zOrder;
        if (worksheet.Charts.Any())
            zOrder = worksheet.Charts.Max(c => c.ZOrder) + 1;
        else
            zOrder = 1;
        ZOrder = zOrder;
        ShapeId = worksheet.Workbook.ShapeIdManager.GetNext();
        RightAngleAxes = true;
    }

    public bool RightAngleAxes { get; set; }
    public IXlsxChart SetRightAngleAxes()
    {
        RightAngleAxes = true;
        return this;
    }
    public IXlsxChart SetRightAngleAxes(bool rightAngleAxes)
    {
        RightAngleAxes = rightAngleAxes;
        return this;
    }

    public XlsxChartType ChartType { get; set; }
    public IXlsxChart SetChartType(XlsxChartType chartType)
    {
        ChartType = chartType;
        return this;
    }

    public XlsxChartTypeCategory ChartTypeCategory
    {
        get
        {
            if (Bar3DCharts.Contains(ChartType))
                return XlsxChartTypeCategory.Bar3D;
            else
                throw new NotImplementedException();

        }
    }

    private readonly HashSet<XlsxChartType> Bar3DCharts = [
        XlsxChartType.BarClustered3D,
        XlsxChartType.BarStacked100Percent3D,
        XlsxChartType.BarStacked3D,
        XlsxChartType.Column3D,
        XlsxChartType.ColumnClustered3D,
        XlsxChartType.ColumnStacked100Percent3D,
        XlsxChartType.ColumnStacked3D
    ];

    public XlsxBarOrientation BarOrientation
    {
        get
        {
            if (HorizontalCharts.Contains(ChartType))
                return XlsxBarOrientation.Horizontal;
            else
                return XlsxBarOrientation.Vertical;
        }
    }

    private readonly HashSet<XlsxChartType> HorizontalCharts = [
        XlsxChartType.BarClustered,
        XlsxChartType.BarClustered3D,
        XlsxChartType.BarStacked,
        XlsxChartType.BarStacked100Percent,
        XlsxChartType.BarStacked100Percent3D,
        XlsxChartType.BarStacked3D,
        XlsxChartType.ConeHorizontalClustered,
        XlsxChartType.ConeHorizontalStacked,
        XlsxChartType.ConeHorizontalStacked100Percent,
        XlsxChartType.CylinderHorizontalClustered,
        XlsxChartType.CylinderHorizontalStacked,
        XlsxChartType.CylinderHorizontalStacked100Percent,
        XlsxChartType.PyramidHorizontalClustered,
        XlsxChartType.PyramidHorizontalStacked,
        XlsxChartType.PyramidHorizontalStacked100Percent
    ];

    public XlsxBarGrouping BarGrouping
    {
        get
        {
            if (ClusteredCharts.Contains(ChartType))
                return XlsxBarGrouping.Clustered;
            else if (PercentCharts.Contains(ChartType))
                return XlsxBarGrouping.Percent;
            else if (StackedCharts.Contains(ChartType))
                return XlsxBarGrouping.Stacked;
            else
                return XlsxBarGrouping.Standard;
        }
    }

    public HashSet<XlsxChartType> ClusteredCharts =
    [
        XlsxChartType.BarClustered,
        XlsxChartType.BarClustered3D,
        XlsxChartType.ColumnClustered,
        XlsxChartType.ColumnClustered3D,
        XlsxChartType.ConeClustered,
        XlsxChartType.ConeHorizontalClustered,
        XlsxChartType.CylinderClustered,
        XlsxChartType.CylinderHorizontalClustered,
        XlsxChartType.PyramidClustered,
        XlsxChartType.PyramidHorizontalClustered
    ];

    public HashSet<XlsxChartType> PercentCharts = [
        XlsxChartType.AreaStacked100Percent,
        XlsxChartType.AreaStacked100Percent3D,
        XlsxChartType.BarStacked100Percent,
        XlsxChartType.BarStacked100Percent3D,
        XlsxChartType.ColumnStacked100Percent,
        XlsxChartType.ColumnStacked100Percent3D,
        XlsxChartType.ConeHorizontalStacked100Percent,
        XlsxChartType.ConeStacked100Percent,
        XlsxChartType.CylinderHorizontalStacked100Percent,
        XlsxChartType.CylinderStacked100Percent,
        XlsxChartType.LineStacked100Percent,
        XlsxChartType.LineWithMarkersStacked100Percent,
        XlsxChartType.PyramidHorizontalStacked100Percent,
        XlsxChartType.PyramidStacked100Percent
    ];

    public HashSet<XlsxChartType> StackedCharts =
    [
        XlsxChartType.AreaStacked,
        XlsxChartType.AreaStacked3D,
        XlsxChartType.BarStacked,
        XlsxChartType.BarStacked3D,
        XlsxChartType.ColumnStacked,
        XlsxChartType.ColumnStacked3D,
        XlsxChartType.ConeHorizontalStacked,
        XlsxChartType.ConeStacked,
        XlsxChartType.CylinderHorizontalStacked,
        XlsxChartType.CylinderStacked,
        XlsxChartType.LineStacked,
        XlsxChartType.LineWithMarkersStacked,
        XlsxChartType.PyramidHorizontalStacked,
        XlsxChartType.PyramidStacked
    ];
}
