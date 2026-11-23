using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Columns;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Rows;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Styles;

/// <summary>
/// API object to modify properties of a cell format of a <see cref="IXlsxFormatContainer"/>.
/// The methods and properties create a modified formats and the formats are registered
/// in the <see cref="XlsxWorkbookStyles"/>.
/// </summary>
internal partial class XlsxCellFormat
{
    private readonly XlsxWorkbook _workbook;
    private readonly Hierarchy _formatValue;

    private XlsxCellFormat(XlsxWorkbook workbook, Hierarchy formatValue)
    {
        _workbook = workbook;
        _formatValue = formatValue;
    }

    internal XlsxNumberCellFormat NumberFormat => new(this);

    internal XlsxFontCellFormat Font => new(this);

    internal XlsxFillCellFormat Fill => new(this);

    internal XlsxBorderCellFormat Border => new(this);

    internal XlsxAlignmentCellFormat Alignment => new(this);

    internal XlsxProtectionCellFormat Protection => new(this);

    internal bool IncludeQuotePrefix
    {
        get => Resolve(static format => format.IncludeQuotePrefix);
        set => ModifyFormat((format, includeQuotePrefix) => format with { IncludeQuotePrefix = includeQuotePrefix }, value);
    }

    /// <summary>
    /// Cell areas in a workbook that should be updated when format is changed, e.g. when we have
    /// a format API object for a row container, the area are all cells of the row. It must be
    /// an area, so we can satisfy the <see cref="IXlsxBorder.OutsideBorder"/> and
    /// <see cref="IXlsxBorder.InsideBorder"/> property setters.
    /// </summary>
    private IReadOnlyList<SheetArea> Areas { get; init; } = [];

    /// <summary>
    /// Formatting is updated for used cells within these areas. Unused cells are ignored.
    /// </summary>
    private IReadOnlyList<SheetArea> UsedAreas { get; init; } = [];

    /// <summary>
    /// Formatting is updated for these columns. This doesn't update cells within the columns, only
    /// the columns themselves. The values are unique columns, sorted by column number in ascending
    /// order.
    /// </summary>
    private XlsxColumnArea[] Columns { get; init; } = [];

    /// <summary>
    /// Formatting is updated for these rows. This doesn't update cells within the rows, only
    /// the rows themselves. The values are unique rows, sorted by row number in ascending order.
    /// </summary>
    private XlsxRowArea[] Rows { get; init; } = [];

    /// <summary>
    /// Formatting is updated for these worksheets. This doesn't update cells within the sheets, only
    /// the sheets and materialized rows and columns of the sheets.
    /// </summary>
    private IReadOnlyList<string> Worksheets { get; init; } = [];

    /// <summary>
    /// Should the formatting be updated for the default format of a workbook (plus cascade to all
    /// formats below, containers and areas).
    /// </summary>
    private bool DefaultFormat { get; init; }

    /// <summary>
    /// A flag indicating API object is for XlsxCells. Unlike other range API objects, the XlsxCells
    /// has a non-standard outside/inside borders behavior.
    /// </summary>
    private bool IsCells { get; init; }

    public override bool Equals(object? obj)
    {
        return obj is IXlsxStyle other && (this as IEquatable<IXlsxStyle>).Equals(other);
    }

    public override int GetHashCode()
    {
        return 0;
    }

    internal static XlsxCellFormat ForCell(XlsxCell cell)
    {
        var workbook = cell.Worksheet.Workbook;
        var sheetName = cell.Worksheet.Name;
        var cellPoint = cell.Point;
        var formatValue = new Hierarchy(workbook, sheetName, cellPoint.Column, cellPoint.Row, cellPoint);
        return new XlsxCellFormat(workbook, formatValue)
        {
            Areas = new[] { new SheetArea(sheetName, new Area(cellPoint)) }
        };
    }

    internal static XlsxCellFormat ForColumn(XlsxColumn column)
    {
        var workbook = column.Worksheet.Workbook;
        var columnArea = column.Area;
        var formatValue = new Hierarchy(workbook, columnArea.Name, columnArea.ColumNumber, null, null);
        return new XlsxCellFormat(workbook, formatValue)
        {
            UsedAreas = new[] { columnArea.Area },
            Columns = [columnArea]
        };
    }

    internal static XlsxCellFormat ForColumns(XlsxWorkbook workbook, XlsxWorksheet? formatValueSheet, IEnumerable<XlsxColumn> columns)
    {
        var columnAreas = columns.Select(x => x.Area).Distinct().OrderBy(x => x.ColumNumber).ToArray();
        var formatValue = new Hierarchy(workbook, formatValueSheet?.Name, null, null, null);
        return new XlsxCellFormat(workbook, formatValue)
        {
            Columns = columnAreas,
            UsedAreas = [.. columnAreas.Select(x => x.Area)]
        };
    }

    internal static XlsxCellFormat ForRow(XlsxRow row)
    {
        var workbook = row.Worksheet.Workbook;
        var rowArea = row.Area;
        var formatValue = new Hierarchy(workbook, rowArea.Name, null, rowArea.RowNumber, null);
        return new XlsxCellFormat(workbook, formatValue)
        {
            UsedAreas = new[] { rowArea.Area },
            Rows = [rowArea]
        };
    }

    internal static XlsxCellFormat ForRows(XlsxWorkbook workbook, XlsxWorksheet? formatValueSheet, IEnumerable<XlsxRow> rows)
    {
        var rowAreas = rows.Select(x => x.Area).Distinct().OrderBy(x => x.RowNumber).ToArray();
        var formatValue = new Hierarchy(workbook, formatValueSheet?.Name, null, null, null);
        return new XlsxCellFormat(workbook, formatValue)
        {
            Rows = rowAreas,
            UsedAreas = [.. rowAreas.Select(x => x.Area)]
        };
    }

    internal static XlsxCellFormat ForWorksheet(XlsxWorksheet worksheet)
    {
        var workbook = worksheet.Workbook;
        var formatValue = new Hierarchy(workbook, worksheet.Name, null, null, null);
        return new XlsxCellFormat(workbook, formatValue)
        {
            UsedAreas = new[] { worksheet.Area },
            Worksheets = new[] { worksheet.Name }
        };
    }

    internal static XlsxCellFormat ForWorkbook(XlsxWorkbook workbook)
    {
        var formatValue = new Hierarchy(workbook, null, null, null, null);
        return new XlsxCellFormat(workbook, formatValue)
        {
            DefaultFormat = true
        };
    }

    internal static XlsxCellFormat ForAreas(XlsxWorkbook workbook, IReadOnlyList<SheetArea> areas, XlsxWorksheet? sheet)
    {
        var formatValue = new Hierarchy(workbook, sheet?.Name, null, null, null);
        return new XlsxCellFormat(workbook, formatValue)
        {
            Areas = areas
        };
    }

    internal static XlsxCellFormat ForCells(XlsxWorkbook workbook, IReadOnlyList<SheetArea> areas, XlsxWorksheet? sheet)
    {
        var formatValue = new Hierarchy(workbook, sheet?.Name, null, null, null);
        return new XlsxCellFormat(workbook, formatValue)
        {
            Areas = areas,
            IsCells = true
        };
    }

    internal static XlsxCellFormat ForRange(XlsxWorksheet sheet, XlsxRangeAddress rangeAddress)
    {
        var workbook = sheet.Workbook;
        var formatValue = new Hierarchy(workbook, sheet.Name, null, null, null);
        return new XlsxCellFormat(workbook, formatValue)
        {
            Areas = new[] { SheetArea.From(rangeAddress) }
        };
    }

    internal static XlsxCellFormat ForTableRows(XlsxWorksheet sheet, SheetArea[] rowAreas)
    {
        var workbook = sheet.Workbook;
        var formatValue = new Hierarchy(workbook, sheet.Name, null, null, null);
        return new XlsxCellFormat(workbook, formatValue)
        {
            Areas = rowAreas
        };
    }

    internal T Resolve<T>(Func<XlsxCellFormatValue, T> selector)
    {
        var format = _formatValue.Resolve();
        return selector(format);
    }

    internal void ModifyFormat<TProperty>(Func<XlsxCellFormatValue, TProperty, XlsxCellFormatValue> modifyFormat, TProperty value)
    {
        var styles = _workbook.Styles;
        Modify(format => styles.GetRegisteredCellFormat(format, cellFormat => modifyFormat(cellFormat, value)));
    }

    // TODO Styles: Move modification methods of each component to the XlsxCellCollection. Modification
    // of component should always update CustomFormat and to make sure that is done, it should be done
    // in a one place.
    internal void ModifyNumberFormat(XlsxNumberFormat numberFormat)
    {
        var styles = _workbook.Styles;
        Modify(format =>
        {
            var modifiedNumberFormat = styles.RegisterNumberFormat(numberFormat);
            var modifiedFormat = styles.GetRegisteredCellFormat(format, cellFormat => cellFormat with
            {
                NumberFormat = modifiedNumberFormat,
                CustomFormat = format.CustomFormat | CellFormatComponents.NumberFormat
            });
            return modifiedFormat;
        });
    }

    internal void ModifyFont<TProperty>(Func<XlsxFontFormatValue, TProperty, XlsxFontFormatValue> modifyFont, TProperty value)
    {
        var styles = _workbook.Styles;
        Modify(format =>
        {
            var modifiedFont = styles.GetRegisteredFontFormat(format.Font, font => modifyFont(font, value));
            var modifiedFormat = styles.GetRegisteredCellFormat(format, cellFormat => cellFormat with
            {
                Font = modifiedFont,
                CustomFormat = format.CustomFormat | CellFormatComponents.Font
            });
            return modifiedFormat;
        });
    }

    internal void ModifyFill<TProperty>(Func<XlsxFillFormatValue, TProperty, XlsxFillFormatValue> modifyFill, TProperty value)
    {
        var styles = _workbook.Styles;
        Modify(format =>
        {
            var modifiedFill = styles.GetRegisteredFillFormat(format.Fill, fill => modifyFill(fill, value));
            var modifiedFormat = styles.GetRegisteredCellFormat(format, cellFormat => cellFormat with
            {
                Fill = modifiedFill,
                CustomFormat = format.CustomFormat | CellFormatComponents.Fill
            });
            return modifiedFormat;
        });
    }

    internal void ModifyBorder<TProperty>(Func<XlsxBorderFormatValue, TProperty, XlsxBorderFormatValue> modifyBorder, TProperty value)
    {
        var styles = _workbook.Styles;
        Modify(GetModifyBorderFunc(border => modifyBorder(border, value), styles));
    }

    internal void ModifyAlignment<TProperty>(Func<XlsxAlignmentFormatValue, TProperty, XlsxAlignmentFormatValue> modifyAlignment, TProperty value)
    {
        var styles = _workbook.Styles;
        Modify(format =>
        {
            var modifiedAlignment = styles.RegisterAlignmentFormat(modifyAlignment(format.Alignment, value));
            var modifiedFormat = styles.GetRegisteredCellFormat(format, cellFormat => cellFormat with
            {
                Alignment = modifiedAlignment,
                CustomFormat = format.CustomFormat | CellFormatComponents.Alignment
            });
            return modifiedFormat;
        });
    }

    internal void ModifyProtection<TProperty>(Func<XlsxProtectionFormatValue, TProperty, XlsxProtectionFormatValue> modifyProtection, TProperty value)
    {
        var styles = _workbook.Styles;
        Modify(format =>
        {
            var modifiedProtection = styles.RegisterProtectionFormat(modifyProtection(format.Protection, value));
            var modifiedFormat = styles.GetRegisteredCellFormat(format, cellFormat => cellFormat with
            {
                Protection = modifiedProtection,
                CustomFormat = format.CustomFormat | CellFormatComponents.Protection
            });
            return modifiedFormat;
        });
    }

    internal void ModifyOuterBorder<TProperty>(Func<XlsxBorderLine, TProperty, XlsxBorderLine> modify, TProperty value)
    {
        var styles = _workbook.Styles;
        if (IsCells)
        {
            var setAll = GetModifyBorderFunc(border => border with
            {
                Left = modify(border.Left, value),
                Top = modify(border.Top, value),
                Right = modify(border.Right, value),
                Bottom = modify(border.Bottom, value),
            }, styles);
            foreach (var (sheetName, area) in Areas)
            {
                if (!_workbook.TryGetWorksheet(sheetName, out XlsxWorksheet worksheet))
                    continue;

                ApplyToAll(area, setAll, worksheet);
            }

            return;
        }

        // Change only top and bottom border of a row. The style is used by non-materialized cells
        // in a row and will be used by non-materialized cells in a row. Same applies to columns.
        var setTopAndBottom = GetModifyBorderFunc(border => border with
        {
            Top = modify(border.Top, value),
            Bottom = modify(border.Bottom, value),
        }, styles);
        ModifyRowsBorder(Rows, setTopAndBottom);

        var setLeftAndRight = GetModifyBorderFunc(border => border with
        {
            Left = modify(border.Left, value),
            Right = modify(border.Right, value),
        }, styles);
        ModifyColumnsBorder(Columns, setLeftAndRight);

        // A normal path for range API object (except XlsxCells). Set outer border to areas.
        // Don't use UsedAreas, they are for columns/rows. Worksheet doesn't have an outer border.
        var setLeft = GetModifyBorderFunc(border => border with { Left = modify(border.Left, value) }, styles);
        var setTop = GetModifyBorderFunc(border => border with { Top = modify(border.Top, value) }, styles);
        var setRight = GetModifyBorderFunc(border => border with { Right = modify(border.Right, value) }, styles);
        var setBottom = GetModifyBorderFunc(border => border with { Bottom = modify(border.Bottom, value) }, styles);
        foreach (var area in Areas)
        {
            if (!_workbook.TryGetWorksheet(area.Name, out XlsxWorksheet worksheet))
                continue;

            var formatResolver = new FormatResolver(worksheet);
            var cellsCollection = worksheet.Internals.CellsCollection;

            // Left side
            var left = area.Area.SliceFromLeft(1);
            cellsCollection.ApplyFormatOnAll(left, setLeft, formatResolver.Resolve);

            // Top side
            var top = area.Area.SliceFromTop(1);
            cellsCollection.ApplyFormatOnAll(top, setTop, formatResolver.Resolve);

            // Right side
            var right = area.Area.SliceFromRight(1);
            cellsCollection.ApplyFormatOnAll(right, setRight, formatResolver.Resolve);

            // Bottom side
            var bottom = area.Area.SliceFromBottom(1);
            cellsCollection.ApplyFormatOnAll(bottom, setBottom, formatResolver.Resolve);
        }
    }

    internal void ModifyInnerBorder<TProperty>(Func<XlsxBorderLine, TProperty, XlsxBorderLine> modify, TProperty value)
    {
        // Shortcut for XlsxCells - it has no inner borders
        if (IsCells)
            return;

        var styles = _workbook.Styles;
        ModifyInsideBordersOfRows(styles, modify, value);
        ModifyInsideBordersOfColumns(styles, modify, value);

        var setLeft = GetModifyBorderFunc(border => border with { Left = modify(border.Left, value) }, styles);
        var setTop = GetModifyBorderFunc(border => border with { Top = modify(border.Top, value) }, styles);
        var setRight = GetModifyBorderFunc(border => border with { Right = modify(border.Right, value) }, styles);
        var setBottom = GetModifyBorderFunc(border => border with { Bottom = modify(border.Bottom, value) }, styles);
        foreach (var (sheetName, area) in Areas)
        {
            if (!_workbook.TryGetWorksheet(sheetName, out XlsxWorksheet worksheet))
                continue;

            var formatResolver = new FormatResolver(worksheet);
            var cellsCollection = worksheet.Internals.CellsCollection;

            // Setting line from both sides is not super useful, but keeps internal state consistent.
            if (area.Width > 1)
            {
                cellsCollection.ApplyFormatOnAll(area.SliceFromLeft(area.Width - 1), setRight, formatResolver.Resolve);
                cellsCollection.ApplyFormatOnAll(area.SliceFromRight(area.Width - 1), setLeft, formatResolver.Resolve);
            }

            if (area.Height > 1)
            {
                cellsCollection.ApplyFormatOnAll(area.SliceFromTop(area.Height - 1), setBottom, formatResolver.Resolve);
                cellsCollection.ApplyFormatOnAll(area.SliceFromBottom(area.Height - 1), setTop, formatResolver.Resolve);
            }
        }
    }

    private static Func<XlsxCellFormatValue, XlsxCellFormatValue> GetModifyBorderFunc(Func<XlsxBorderFormatValue, XlsxBorderFormatValue> modifyBorder, XlsxWorkbookStyles styles)
    {
        return format =>
        {
            var modifiedBorder = styles.GetRegisteredBorderFormat(format.Border, border =>
            {
                var modified = modifyBorder(border);

                // Per original behavior, the non-visible border can't hold color state, e.g. when
                // a border is set to from Thin to None and later changed back to Thick, it
                // shouldn't remember the original color.
                // That is not how Excel behaves and it makes everything harder (e.g. user can't
                // set the border color first and then border style), but it is what it is.
                if (!modified.Left.IsVisible)
                    modified = modified with { Left = XlsxBorderLine.None };

                if (!modified.Top.IsVisible)
                    modified = modified with { Top = XlsxBorderLine.None };

                if (!modified.Right.IsVisible)
                    modified = modified with { Right = XlsxBorderLine.None };

                if (!modified.Bottom.IsVisible)
                    modified = modified with { Bottom = XlsxBorderLine.None };

                if (!modified.Diagonal.IsVisible)
                    modified = modified with { Diagonal = XlsxBorderLine.None };

                return modified;
            });
            var modifiedFormat = styles.GetRegisteredCellFormat(format, cellFormat => cellFormat with
            {
                Border = modifiedBorder,
                CustomFormat = format.CustomFormat | CellFormatComponents.Border
            });
            return modifiedFormat;
        };
    }

    private void ModifyInsideBordersOfRows<TProperty>(XlsxWorkbookStyles styles, Func<XlsxBorderLine, TProperty, XlsxBorderLine> modify, TProperty value)
    {
        // For a single row, only the left are right border are counted as "inside". The top and bottom border touch the outside.
        var setLeftAndRight = GetModifyBorderFunc(border => border with
        {
            Left = modify(border.Left, value),
            Right = modify(border.Right, value),
        }, styles);

        // For multi-row rowspan, there are three different patterns:
        // Multi-row rowspan - top row
        var setLeftRightBottom = GetModifyBorderFunc(border => border with
        {
            Left = modify(border.Left, value),
            Right = modify(border.Right, value),
            Bottom = modify(border.Bottom, value),
        }, styles);

        // Multi-row rowspan - center rows. There isn't a center row in 2-row rowspan
        var setAll = GetModifyBorderFunc(border => border with
        {
            Left = modify(border.Left, value),
            Top = modify(border.Top, value),
            Right = modify(border.Right, value),
            Bottom = modify(border.Bottom, value),
        }, styles);

        // Multi-row rowspan - bottom row
        var setLeftTopRight = GetModifyBorderFunc(border => border with
        {
            Left = modify(border.Left, value),
            Top = modify(border.Top, value),
            Right = modify(border.Right, value),
        }, styles);

        // Set border for each rowspan
        for (var i = 0; i < Rows.Length; ++i)
        {
            // Find rowspan as a sequence of consecutive rows 
            var startIndex = i;
            var endIndex = i;
            while (endIndex + 1 < Rows.Length && Rows[endIndex + 1].RowNumber == Rows[endIndex].RowNumber + 1)
            {
                endIndex += 1;
            }

            i = endIndex;
            var rowSpanHeight = endIndex - startIndex + 1;
            if (rowSpanHeight > 1)
            {
                ModifyRowsBorder(Rows.AsSpan(startIndex, 1), setLeftRightBottom);
                ModifyRowsBorder(Rows.AsSpan(startIndex + 1, rowSpanHeight - 2), setAll);
                ModifyRowsBorder(Rows.AsSpan(endIndex, 1), setLeftTopRight);
            }
            else
            {
                ModifyRowsBorder(Rows.AsSpan(startIndex, 1), setLeftAndRight);
            }
        }
    }

    private void ModifyRowsBorder(ReadOnlySpan<XlsxRowArea> rows, Func<XlsxCellFormatValue, XlsxCellFormatValue> modifyBorder)
    {
        foreach (var rowArea in rows)
        {
            if (!_workbook.TryGetWorksheet(rowArea.Name, out XlsxWorksheet worksheet))
                continue;

            // Row style is used by non-materialized cells in a row...
            var row = worksheet.Row(rowArea.RowNumber);
            ApplyColRowFormat(row, modifyBorder, worksheet);

            // ... and materialized cells in a row have format explicitly set.
            var formatResolver = new FormatResolver(worksheet);
            var cellsCollection = worksheet.Internals.CellsCollection;
            cellsCollection.ApplyFormatOnUsed(rowArea.Area.Area, modifyBorder, formatResolver.Resolve);
        }
    }

    private void ModifyInsideBordersOfColumns<TProperty>(XlsxWorkbookStyles styles, Func<XlsxBorderLine, TProperty, XlsxBorderLine> modify, TProperty value)
    {
        // For a single column, only the top are bottom border are counted as "inside". The left and right border touch the outside.
        var setTopAndBottom = GetModifyBorderFunc(border => border with
        {
            Top = modify(border.Top, value),
            Bottom = modify(border.Bottom, value)
        }, styles);

        // For multi-column colspan, there are three different patterns:
        // Multi-column colspan - left column
        var setTopRightBottom = GetModifyBorderFunc(border => border with
        {
            Top = modify(border.Top, value),
            Right = modify(border.Right, value),
            Bottom = modify(border.Bottom, value),
        }, styles);

        // Multi-column colspan - center columns. There isn't a center column in 2-column colspan
        var setAll = GetModifyBorderFunc(border => border with
        {
            Left = modify(border.Left, value),
            Top = modify(border.Top, value),
            Right = modify(border.Right, value),
            Bottom = modify(border.Bottom, value),
        }, styles);

        // Multi-column colspan - right column
        var setLeftTopBottom = GetModifyBorderFunc(border => border with
        {
            Left = modify(border.Left, value),
            Top = modify(border.Top, value),
            Bottom = modify(border.Bottom, value),
        }, styles);

        // Set border for each colspan
        for (var i = 0; i < Columns.Length; ++i)
        {
            // Find colspan as a sequence of consecutive columns 
            var startIndex = i;
            var endIndex = i;
            while (endIndex + 1 < Columns.Length && Columns[endIndex + 1].ColumNumber == Columns[endIndex].ColumNumber + 1)
            {
                endIndex += 1;
            }

            i = endIndex;
            var colspanWidth = endIndex - startIndex + 1;
            if (colspanWidth > 1)
            {
                ModifyColumnsBorder(Columns.AsSpan(startIndex, 1), setTopRightBottom);
                ModifyColumnsBorder(Columns.AsSpan(startIndex + 1, colspanWidth - 2), setAll);
                ModifyColumnsBorder(Columns.AsSpan(endIndex, 1), setLeftTopBottom);
            }
            else
            {
                ModifyColumnsBorder(Columns.AsSpan(startIndex, 1), setTopAndBottom);
            }
        }
    }

    private void ModifyColumnsBorder(ReadOnlySpan<XlsxColumnArea> columns, Func<XlsxCellFormatValue, XlsxCellFormatValue> modifyBorder)
    {
        foreach (var columnArea in columns)
        {
            if (!_workbook.TryGetWorksheet(columnArea.Name, out XlsxWorksheet worksheet))
                continue;

            // Column style is used by non-materialized cells in a column...
            var column = worksheet.Column(columnArea.ColumNumber);
            ApplyColRowFormat(column, modifyBorder, worksheet);

            // ... and materialized cells in a column have format explicitly set.
            var formatResolver = new FormatResolver(worksheet);
            var cellsCollection = worksheet.Internals.CellsCollection;
            cellsCollection.ApplyFormatOnUsed(columnArea.Area.Area, modifyBorder, formatResolver.Resolve);
        }
    }

    private void Modify(Func<XlsxCellFormatValue, XlsxCellFormatValue> modifyFormat)
    {
        // TODO Styles: Deal with cross points
        var styles = _workbook.Styles;
        if (DefaultFormat)
        {
            styles.DefaultFormat = modifyFormat(styles.DefaultFormat);
            foreach (var worksheet in _workbook.WorksheetsInternal)
            {
                ApplyToWorksheet(worksheet, modifyFormat, styles);
            }
        }

        foreach (var sheetName in Worksheets)
        {
            if (!_workbook.TryGetWorksheet(sheetName, out XlsxWorksheet worksheet))
                continue;

            ApplyToWorksheet(worksheet, modifyFormat, styles);
        }

        foreach (var columnArea in Columns)
        {
            if (!_workbook.TryGetWorksheet(columnArea.Name, out XlsxWorksheet worksheet))
                continue;

            var column = worksheet.Column(columnArea.ColumNumber);
            ApplyColRowFormat(column, modifyFormat, worksheet);
        }

        foreach (var rowArea in Rows)
        {
            if (!_workbook.TryGetWorksheet(rowArea.Name, out XlsxWorksheet worksheet))
                continue;

            var row = worksheet.Row(rowArea.RowNumber);
            ApplyColRowFormat(row, modifyFormat, worksheet);
        }

        foreach (var (sheetName, area) in UsedAreas)
        {
            if (!_workbook.TryGetWorksheet(sheetName, out XlsxWorksheet worksheet))
                continue;

            ApplyToUsed(area, modifyFormat, worksheet);
        }

        foreach (var (sheetName, area) in Areas)
        {
            if (!_workbook.TryGetWorksheet(sheetName, out XlsxWorksheet worksheet))
                continue;

            ApplyToAll(area, modifyFormat, worksheet);
        }
    }

    private static void ApplyToWorksheet(XlsxWorksheet worksheet, Func<XlsxCellFormatValue, XlsxCellFormatValue> modifyFormat, XlsxWorkbookStyles styles)
    {
        var originalFormat = worksheet.FormatValue ?? styles.DefaultFormat;
        var modifiedFormat = modifyFormat(originalFormat);
        worksheet.FormatValue = modifiedFormat;

        var columns = worksheet.Internals.ColumnsCollection.Values;
        foreach (var column in columns)
            ApplyColRowFormat(column, modifyFormat, worksheet);

        var rows = worksheet.Internals.RowsCollection.Values;
        foreach (var row in rows)
            ApplyColRowFormat(row, modifyFormat, worksheet);

        ApplyToUsed(Area.Full, modifyFormat, worksheet);
    }

    private static void ApplyColRowFormat(IXlsxFormatContainer rowOrCol, Func<XlsxCellFormatValue, XlsxCellFormatValue> modifyFormat, XlsxWorksheet worksheet)
    {
        if (rowOrCol.FormatValue is not { } originalFormat)
            originalFormat = worksheet.FormatValue ?? worksheet.Workbook.Styles.DefaultFormat;

        rowOrCol.FormatValue = modifyFormat(originalFormat);
    }

    private static void ApplyToUsed(Area area, Func<XlsxCellFormatValue, XlsxCellFormatValue> modifyFormat, XlsxWorksheet worksheet)
    {
        var formatResolver = new FormatResolver(worksheet);
        var cellsCollection = worksheet.Internals.CellsCollection;
        cellsCollection.ApplyFormatOnUsed(area, modifyFormat, formatResolver.Resolve);
    }

    private static void ApplyToAll(Area area, Func<XlsxCellFormatValue, XlsxCellFormatValue> modifyFormat, XlsxWorksheet worksheet)
    {
        var formatResolver = new FormatResolver(worksheet);
        var cellsCollection = worksheet.Internals.CellsCollection;
        cellsCollection.ApplyFormatOnAll(area, modifyFormat, formatResolver.Resolve);
    }

    /// <summary>
    /// A format value resolution hierarchy for a range API object. Each range API type needs
    /// to set proper fallbacks through ctor.
    /// </summary>
    private readonly record struct Hierarchy
    {
        private readonly XlsxWorkbook _workbook;
        private readonly string? _sheetName;
        private readonly int? _columnNumber;
        private readonly int? _rowNumber;
        private readonly Point? _point;

        public Hierarchy(XlsxWorkbook workbook, string? sheetName, int? columnNumber, int? rowNumber, Point? point)
        {
            _workbook = workbook;
            _sheetName = sheetName;
            _columnNumber = columnNumber;
            _rowNumber = rowNumber;
            _point = point;
        }

        private XlsxCellFormatValue DefaultFormat => _workbook.Styles.DefaultFormat;

        internal XlsxCellFormatValue Resolve()
        {
            var isForWorkbook = _sheetName is null;
            if (isForWorkbook)
                return DefaultFormat;

            // First, make sure the sheet exists
            if (!_workbook.TryGetWorksheet(_sheetName, out XlsxWorksheet sheet))
                return DefaultFormat;

            if (_point is { } point)
            {
                var formatSlice = sheet.Internals.CellsCollection.FormatSlice;
                var cellFormat = formatSlice.GetFormat(point);
                if (cellFormat is not null)
                    return cellFormat;
            }

            if (_rowNumber is { } rowNumber)
            {
                var rowsCollection = sheet.Internals.RowsCollection;
                if (rowsCollection.TryGetValue(rowNumber, out var row) &&
                    row.FormatValue is { } rowFormat)
                {
                    return rowFormat;
                }
            }

            if (_columnNumber is { } columnNumber)
            {
                var columnsCollection = sheet.Internals.ColumnsCollection;
                if (columnsCollection.TryGetValue(columnNumber, out var column) &&
                    column.FormatValue is { } columnFormat)
                {
                    return columnFormat;
                }
            }

            if (sheet.FormatValue is { } sheetFormat)
                return sheetFormat;

            return DefaultFormat;
        }
    }
}
