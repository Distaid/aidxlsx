#nullable disable


namespace Aid.Xlsx.Model.Drawings.Charts;

public enum XlsxChartType
{
    Area,
    Area3D,
    AreaStacked,
    AreaStacked100Percent,
    AreaStacked100Percent3D,
    AreaStacked3D,
    BarClustered,
    BarClustered3D,
    BarStacked,
    BarStacked100Percent,
    BarStacked100Percent3D,
    BarStacked3D,
    Bubble,
    Bubble3D,
    Column3D,
    ColumnClustered,
    ColumnClustered3D,
    ColumnStacked,
    ColumnStacked100Percent,
    ColumnStacked100Percent3D,
    ColumnStacked3D,
    Cone,
    ConeClustered,
    ConeHorizontalClustered,
    ConeHorizontalStacked,
    ConeHorizontalStacked100Percent,
    ConeStacked,
    ConeStacked100Percent,
    Cylinder,
    CylinderClustered,
    CylinderHorizontalClustered,
    CylinderHorizontalStacked,
    CylinderHorizontalStacked100Percent,
    CylinderStacked,
    CylinderStacked100Percent,
    Doughnut,
    DoughnutExploded,
    Line,
    Line3D,
    LineStacked,
    LineStacked100Percent,
    LineWithMarkers,
    LineWithMarkersStacked,
    LineWithMarkersStacked100Percent,
    Pie,
    Pie3D,
    PieExploded,
    PieExploded3D,
    PieToBar,
    PieToPie,
    Pyramid,
    PyramidClustered,
    PyramidHorizontalClustered,
    PyramidHorizontalStacked,
    PyramidHorizontalStacked100Percent,
    PyramidStacked,
    PyramidStacked100Percent,
    Radar,
    RadarFilled,
    RadarWithMarkers,
    StockHighLowClose,
    StockOpenHighLowClose,
    StockVolumeHighLowClose,
    StockVolumeOpenHighLowClose,
    Surface,
    SurfaceContour,
    SurfaceContourWireframe,
    SurfaceWireframe,
    XYScatterMarkers,
    XYScatterSmoothLinesNoMarkers,
    XYScatterSmoothLinesWithMarkers,
    XYScatterStraightLinesNoMarkers,
    XYScatterStraightLinesWithMarkers
}
public interface IXlsxChart : IXlsxDrawing<IXlsxChart>
{
    bool RightAngleAxes { get; set; }
    IXlsxChart SetRightAngleAxes();
    IXlsxChart SetRightAngleAxes(bool rightAngleAxes);

    XlsxChartType ChartType { get; set; }
    IXlsxChart SetChartType(XlsxChartType chartType);

}
