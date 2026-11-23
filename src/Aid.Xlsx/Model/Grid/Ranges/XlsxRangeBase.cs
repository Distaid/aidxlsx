#nullable disable

using System.Globalization;
using System.Text;
using Aid.Xlsx.Calculation.Visitors;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.AutoFilters;
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.DataValidation;
using Aid.Xlsx.Model.DefinedNames;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Columns;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges.Sort;
using Aid.Xlsx.Model.Grid.Rows;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Grid.Ranges;

internal abstract class XlsxRangeBase(XlsxRangeAddress rangeAddress) : IXlsxRangeBase
{
    private XlsxSortElements _sortRows;
    private XlsxSortElements _sortColumns;

    /// <summary>
    /// Get format API object tailored to the range type.
    /// </summary>
    internal abstract XlsxCellFormat Format { get; }

    protected virtual void OnRangeAddressChanged(XlsxRangeAddress oldAddress, XlsxRangeAddress newAddress)
    {
        Worksheet.RelocateRange(RangeType, oldAddress, newAddress);
    }

    #region Public properties

    private XlsxRangeAddress _rangeAddress = rangeAddress;

    public XlsxRangeAddress RangeAddress
    {
        get { return _rangeAddress; }
        protected set
        {
            if (_rangeAddress != value)
            {
                var oldAddress = _rangeAddress;
                _rangeAddress = value;
                OnRangeAddressChanged(oldAddress, _rangeAddress);
            }
        }
    }

    public XlsxWorksheet Worksheet
    {
        get { return RangeAddress.Worksheet; }
    }

    internal Area SheetRange
    {
        get
        {
            if (!RangeAddress.IsValid)
                throw new InvalidOperationException("Range address is invalid.");

            return Area.FromRangeAddress(RangeAddress);
        }
    }

    IXlsxDataValidation IXlsxRangeBase.CreateDataValidation()
    {
        return CreateDataValidation();
    }

    internal XlsxDataValidation CreateDataValidation()
    {
        return Worksheet.DataValidations.Create(SheetRange);
    }

    public IXlsxDataValidation GetDataValidation()
    {
        Worksheet.DataValidations.TryGet(RangeAddress, out var existingDataValidation);
        return existingDataValidation;
    }

    #region IXlsxRangeBase Members

    public IXlsxStyle Style
    {
        get => Format;
        set => Format.SetStyle(value);
    }

    IXlsxRangeAddress IXlsxAddressable.RangeAddress
    {
        get { return RangeAddress; }
    }

    IXlsxWorksheet IXlsxRangeBase.Worksheet
    {
        get { return RangeAddress.Worksheet; }
    }

    public String FormulaA1
    {
        set
        {
            Cells().ForEach(c =>
                                {
                                    c.FormulaA1 = value;
                                    c.FormulaReference = RangeAddress;
                                });
        }
    }

    public String FormulaArrayA1
    {
        set
        {
            var range = Area.FromRangeAddress(RangeAddress);
            if (Worksheet.MergedRanges.Any(mr => mr.Intersects(this)))
                throw new InvalidOperationException("Can't create array function over a merged range.");

            if (Worksheet.Tables.Any<XlsxTable>(t => t.Intersects(this)))
                throw new InvalidOperationException("Can't create array function over a table.");

            if (Cells(false).Any<XlsxCell>(c => c.HasArrayFormula && !RangeAddress.ContainsWhole(c.FormulaReference)))
                throw new InvalidOperationException("Can't create array function that partially covers another array function.");

            var formula = value.TrimFormulaEqual();
            var fixedFunctionsFormula = FormulaTransformation.FixFutureFunctions(formula, Worksheet.Name, SheetRange.FirstPoint);
            var arrayFormula = XlsxCellFormula.Array(fixedFunctionsFormula, range, false);

            var formulaSlice = Worksheet.Internals.CellsCollection.FormulaSlice;
            formulaSlice.SetArray(range, arrayFormula);

            // If formula evaluates to a text, it is stored directly in a worksheet, not in SST. Thus
            // when the switch to formula happens, disable shared string and enable when formula is removed.
            var valueSlice = Worksheet.Internals.CellsCollection.ValueSlice;
            for (var row = range.TopRow; row <= range.BottomRow; ++row)
            {
                for (var col = range.LeftColumn; col <= range.RightColumn; ++col)
                {
                    valueSlice.SetShareString(new Point(row, col), false);
                }
            }

            // Formula is shared across all cells, so it's enough to invalidate master cell
            var masterCell = FirstCell();
            masterCell.InvalidateFormula();
        }
    }

    public String FormulaR1C1
    {
        set
        {
            Cells().ForEach(c =>
            {
                c.FormulaR1C1 = value;
                c.FormulaReference = RangeAddress;
            });
        }
    }

    public Boolean ShareString
    {
        set { Cells().ForEach(c => c.ShareString = value); }
    }

    public XlsxCellValue Value
    {
        set { Cells().ForEach(c => c.Value = value); }
    }

    #endregion IXlsxRangeBase Members

    #endregion Public properties

    #region IXlsxRangeBase Members

    IXlsxCells IXlsxRangeBase.Cells(String cells) => Cells(cells);

    IXlsxCells IXlsxRangeBase.Cells(Boolean usedCellsOnly) => Cells(usedCellsOnly);

    IXlsxCells IXlsxRangeBase.Cells(Boolean usedCellsOnly, XlsxCellsUsedOptions options) => Cells(usedCellsOnly, options);

    IXlsxCells IXlsxRangeBase.CellsUsed() => CellsUsed();

    IXlsxCell IXlsxRangeBase.FirstCell()
    {
        return FirstCell();
    }

    IXlsxCell IXlsxRangeBase.LastCell()
    {
        return LastCell();
    }

    IXlsxCell IXlsxRangeBase.FirstCellUsed()
    {
        return FirstCellUsed(XlsxCellsUsedOptions.AllContents);
    }

    IXlsxCell IXlsxRangeBase.FirstCellUsed(XlsxCellsUsedOptions options)
    {
        return FirstCellUsed(options, null);
    }

    IXlsxCell IXlsxRangeBase.FirstCellUsed(Func<IXlsxCell, Boolean> predicate)
    {
        return FirstCellUsed(predicate);
    }

    IXlsxCell IXlsxRangeBase.FirstCellUsed(XlsxCellsUsedOptions options, Func<IXlsxCell, Boolean> predicate)
    {
        return FirstCellUsed(options, predicate);
    }

    IXlsxCell IXlsxRangeBase.LastCellUsed()
    {
        return LastCellUsed(XlsxCellsUsedOptions.AllContents);
    }

    IXlsxCell IXlsxRangeBase.LastCellUsed(XlsxCellsUsedOptions options)
    {
        return LastCellUsed(options, null);
    }

    IXlsxCell IXlsxRangeBase.LastCellUsed(Func<IXlsxCell, Boolean> predicate)
    {
        return LastCellUsed(predicate);
    }

    IXlsxCell IXlsxRangeBase.LastCellUsed(XlsxCellsUsedOptions options, Func<IXlsxCell, Boolean> predicate)
    {
        return LastCellUsed(options, predicate);
    }

    public virtual IXlsxCells Cells()
    {
        return Cells(false);
    }

    public virtual XlsxCells Cells(Boolean usedCellsOnly)
    {
        return Cells(usedCellsOnly, XlsxCellsUsedOptions.AllContents);
    }

    public XlsxCells Cells(Boolean usedCellsOnly, XlsxCellsUsedOptions options)
    {
        var cells = new XlsxCells(Worksheet, usedCellsOnly, options) { RangeAddress };
        return cells;
    }

    public virtual XlsxCells Cells(String cells)
    {
        return Ranges(cells).Cells();
    }

    public IXlsxCells Cells(Func<IXlsxCell, Boolean> predicate)
    {
        var cells = new XlsxCells(Worksheet, false, XlsxCellsUsedOptions.AllContents, predicate) { RangeAddress };
        return cells;
    }

    public XlsxCells CellsUsed()
    {
        return Cells(true);
    }

    public IXlsxRange Merge()
    {
        return Merge(true);
    }

    public IXlsxRange Merge(Boolean checkIntersect)
    {
        if (RangeAddress.FirstAddress.Equals(RangeAddress.LastAddress))
            return Worksheet.Range(RangeAddress);

        var asRange = AsRange();

        if (checkIntersect)
        {
            var intersectedMergedRanges =
                Worksheet.Internals.MergedRanges.GetIntersectedRanges(RangeAddress).ToList();
            foreach (var intersectedMergedRange in intersectedMergedRanges)
            {
                Worksheet.Internals.MergedRanges.Remove(intersectedMergedRange);
            }

            var firstCell = FirstCell();
            var cellsUsed =
                CellsUsedInternal(XlsxCellsUsedOptions.All & ~XlsxCellsUsedOptions.MergedRanges, c => c.Point != firstCell.Point).ToList<XlsxCell>();
            cellsUsed.ForEach(c => c.Clear(XlsxClearOptions.All
                                           & ~XlsxClearOptions.MergedRanges
                                           & ~XlsxClearOptions.NormalFormats));

            // When a range is merged, remaining cells of the area take on the format of the first cell
            if (firstCell.FormatValue is null)
            {
                Worksheet.Internals.CellsCollection.FormatSlice.Clear(SheetRange);
            }
            else
            {
                // Merging removes borders that are not consistent across the whole border, even on the first cell
                var area = SheetRange;
                var leftBorder = GetVerticalBorder(area.LeftColumn, area.TopRow, area.BottomRow, static b => b.Left);
                var topBorder = GetHorizontalBorder(area.TopRow, area.LeftColumn, area.RightColumn, static b => b.Top);
                var rightBorder = GetVerticalBorder(area.RightColumn, area.TopRow, area.BottomRow, static b => b.Right);
                var bottomBorder = GetHorizontalBorder(area.BottomRow, area.LeftColumn, area.RightColumn, static b => b.Bottom);

                var cellsCollection = Worksheet.Internals.CellsCollection;
                var borderlessFormat = Worksheet.Workbook.Styles.GetModifiedFormat(firstCell.FormatValue, _ => XlsxBorderFormatValue.None);
                cellsCollection.FormatSlice.SetAll(SheetRange, borderlessFormat);

                if (leftBorder is not null && leftBorder.Value.IsVisible)
                    cellsCollection.ApplyFormatOnAll(SheetRange.SliceFromLeft(1), b => b with { Left = leftBorder.Value });

                if (topBorder is not null && topBorder.Value.IsVisible)
                    cellsCollection.ApplyFormatOnAll(SheetRange.SliceFromTop(1), b => b with { Top = topBorder.Value });

                if (rightBorder is not null && rightBorder.Value.IsVisible)
                    cellsCollection.ApplyFormatOnAll(SheetRange.SliceFromRight(1), b => b with { Right = rightBorder.Value });

                if (bottomBorder is not null && bottomBorder.Value.IsVisible)
                    cellsCollection.ApplyFormatOnAll(SheetRange.SliceFromBottom(1), b => b with { Bottom = bottomBorder.Value });
            }
        }

        Worksheet.Internals.MergedRanges.Add(asRange);
        return asRange;
    }

    private XlsxBorderLine? GetHorizontalBorder(int row, int minColumn, int maxColumn, Func<XlsxBorderFormatValue, XlsxBorderLine> borderSide)
    {
        var initialSide = borderSide(Worksheet.GetStyleValue(new Point(row, minColumn)).Border);
        for (var column = minColumn + 1; column <= maxColumn; ++column)
        {
            var currentSide = borderSide(Worksheet.GetStyleValue(new Point(row, column)).Border);
            if (currentSide != initialSide)
                return null;
        }

        return initialSide;
    }

    private XlsxBorderLine? GetVerticalBorder(int column, int minRow, int maxRow, Func<XlsxBorderFormatValue, XlsxBorderLine> borderSide)
    {
        var initialSide = borderSide(Worksheet.GetStyleValue(new Point(minRow, column)).Border);
        for (var row = minRow + 1; row <= maxRow; ++row)
        {
            var currentSide = borderSide(Worksheet.GetStyleValue(new Point(row, column)).Border);
            if (currentSide != initialSide)
                return null;
        }

        return initialSide;
    }

    public IXlsxRange Unmerge()
    {
        string tAddress = RangeAddress.ToString();
        var asRange = AsRange();
        if (Worksheet.Internals.MergedRanges.Select<XlsxRange, string>(m => m.RangeAddress.ToString()).Any(mAddress => mAddress == tAddress))
            Worksheet.Internals.MergedRanges.Remove(asRange);

        return asRange;
    }

    public IXlsxRangeBase Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
    {
        var cellClearOptions = clearOptions
                & ~XlsxClearOptions.ConditionalFormats
                & ~XlsxClearOptions.DataValidation
                & ~XlsxClearOptions.MergedRanges
                & ~XlsxClearOptions.Sparklines;
        var cellUsedOptions = cellClearOptions.ToCellsUsedOptions();
        foreach (var cell in CellsUsedInternal(cellUsedOptions))
        {
            // We'll clear the conditional formatting, data validations, sparklines
            // and merged ranges later down.
            cell.Clear(cellClearOptions, true);
        }

        if (clearOptions.HasFlag(XlsxClearOptions.ConditionalFormats))
            Worksheet.ConditionalFormats.Clear(SheetRange);

        if (clearOptions.HasFlag(XlsxClearOptions.DataValidation))
        {
            var validation = CreateDataValidation();
            Worksheet.DataValidations.Delete(validation);
        }

        if (clearOptions.HasFlag(XlsxClearOptions.MergedRanges))
            ClearMerged();

        if (clearOptions.HasFlag(XlsxClearOptions.Sparklines))
            RemoveSparklines();

        if (clearOptions == XlsxClearOptions.All)
        {
            Worksheet.Internals.CellsCollection.Clear(Area.FromRangeAddress(RangeAddress));
        }
        return this;
    }

    public IXlsxRangeBase Relative(IXlsxRangeBase sourceBaseRange, IXlsxRangeBase targetBaseRange)
    {
        var xlSourceBaseRangeAddress = (XlsxRangeAddress)sourceBaseRange.RangeAddress;
        var xlTargetBaseRangeAddress = (XlsxRangeAddress)targetBaseRange.RangeAddress;
        var xlRangeAddress = this.RangeAddress.Relative(in xlSourceBaseRangeAddress, in xlTargetBaseRangeAddress);

        return ((XlsxRangeBase)targetBaseRange).Range(in xlRangeAddress);
    }

    internal void RemoveSparklines()
    {
        Worksheet.SparklineGroups.GetSparklines(this).ToList()
            .ForEach(sl => Worksheet.SparklineGroups.Remove(sl.Location));
    }

    public void DeleteComments()
    {
        Cells().DeleteComments();
    }

    public bool Contains(String rangeAddress)
    {
        string addressToUse = rangeAddress.Contains('!')
                                  ? rangeAddress[(rangeAddress.IndexOf('!') + 1)..]
                                  : rangeAddress;

        XlsxAddress firstAddress;
        XlsxAddress lastAddress;
        if (addressToUse.Contains(':'))
        {
            var arrRange = addressToUse.Split(':');
            firstAddress = XlsxAddress.Create(Worksheet, arrRange[0]);
            lastAddress = XlsxAddress.Create(Worksheet, arrRange[1]);
        }
        else
        {
            firstAddress = XlsxAddress.Create(Worksheet, addressToUse);
            lastAddress = XlsxAddress.Create(Worksheet, addressToUse);
        }
        return Contains(firstAddress, lastAddress);
    }

    public bool Contains(IXlsxRangeBase range)
    {
        return Contains((XlsxAddress)range.RangeAddress.FirstAddress, (XlsxAddress)range.RangeAddress.LastAddress);
    }

    public bool Intersects(string rangeAddress)
    {
        return Intersects(Worksheet.Range(rangeAddress));
    }

    public bool Intersects(IXlsxRangeBase range)
    {
        if (!range.RangeAddress.IsValid || !RangeAddress.IsValid)
            return false;
        var ma = range.RangeAddress;
        var ra = RangeAddress;
        return ra.Intersects(ma);
    }

    IXlsxRange IXlsxRangeBase.AsRange()
    {
        return AsRange();
    }

    public virtual XlsxRange AsRange()
    {
        return Worksheet.Range(RangeAddress);
    }

    public IXlsxRange AddToNamed(String name)
    {
        return AddToNamed(name, XlsxScope.Workbook);
    }

    public IXlsxRange AddToNamed(String name, XlsxScope scope)
    {
        return AddToNamed(name, scope, null);
    }

    public IXlsxRange AddToNamed(String name, XlsxScope scope, String comment)
    {
        var definedNames = scope == XlsxScope.Workbook
                              ? Worksheet.Workbook.DefinedNamesInternal
                              : Worksheet.DefinedNames;

        if (definedNames.TryGetScopedValue(name, out var definedName))
            definedName.Add(RangeAddress.ToStringFixed(XlsxReferenceStyle.A1, true));
        else
            definedNames.Add(name, RangeAddress.ToStringFixed(XlsxReferenceStyle.A1, true), comment);

        return AsRange();
    }

    public IXlsxRangeBase SetValue(XlsxCellValue value)
    {
        Cells().ForEach(c => c.SetValue(value));
        return this;
    }

    public Boolean IsMerged()
    {
        return Cells().Any(c => c.IsMerged());
    }

    public virtual Boolean IsEmpty()
    {
        return !CellsUsed().Any<XlsxCell>() || CellsUsed().Any<XlsxCell>(c => c.IsEmpty());
    }

    public virtual Boolean IsEmpty(XlsxCellsUsedOptions options)
    {
        foreach (var cell in CellsUsed(options))
        {
            if (!cell.IsEmpty(options))
            {
                return false;
            }
        }
        return true;
    }

    public virtual Boolean IsEntireRow()
    {
        return RangeAddress.IsEntireRow();
    }

    public virtual Boolean IsEntireColumn()
    {
        return RangeAddress.IsEntireColumn();
    }

    public Boolean IsEntireSheet()
    {
        return RangeAddress.IsEntireSheet();
    }

    #endregion IXlsxRangeBase Members

    public IXlsxCells Search(String searchText, CompareOptions compareOptions = CompareOptions.Ordinal, Boolean searchFormulae = false)
    {
        var culture = CultureInfo.CurrentCulture;
        return CellsUsed(XlsxCellsUsedOptions.AllContents, c =>
        {
            try
            {
                if (searchFormulae)
                    return c.HasFormula
                           && culture.CompareInfo.IndexOf(c.FormulaA1, searchText, compareOptions) >= 0
                           || culture.CompareInfo.IndexOf(c.Value.ToString(CultureInfo.CurrentCulture), searchText, compareOptions) >= 0;
                else
                    return culture.CompareInfo.IndexOf(c.Value.ToString(CultureInfo.CurrentCulture), searchText, compareOptions) >= 0;
            }
            catch
            {
                return false;
            }
        });
    }

    internal XlsxCell FirstCell()
    {
        return Cell(1, 1);
    }

    internal XlsxCell LastCell()
    {
        return Cell(RowCount(), ColumnCount());
    }

    internal XlsxCell FirstCellUsed()
    {
        return FirstCellUsed(XlsxCellsUsedOptions.AllContents, predicate: null);
    }

    internal XlsxCell FirstCellUsed(Func<IXlsxCell, Boolean> predicate)
    {
        return FirstCellUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    internal XlsxCell FirstCellUsed(XlsxCellsUsedOptions options, Func<IXlsxCell, Boolean> predicate = null)
    {
        var cellsUsed = CellsUsedInternal(options, r => r.FirstCell(), predicate).ToList();

        if (cellsUsed.Count == 0)
            return null;

        var firstRow = cellsUsed.Min(c => c.Address.RowNumber);
        var firstColumn = cellsUsed.Min(c => c.Address.ColumnNumber);

        if (firstRow < RangeAddress.FirstAddress.RowNumber)
            firstRow = RangeAddress.FirstAddress.RowNumber;

        if (firstColumn < RangeAddress.FirstAddress.ColumnNumber)
            firstColumn = RangeAddress.FirstAddress.ColumnNumber;

        return Worksheet.Cell(firstRow, firstColumn);
    }

    internal XlsxCell LastCellUsed()
    {
        return LastCellUsed(XlsxCellsUsedOptions.AllContents, predicate: null);
    }

    internal XlsxCell LastCellUsed(Func<IXlsxCell, Boolean> predicate)
    {
        return LastCellUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    internal XlsxCell LastCellUsed(XlsxCellsUsedOptions options, Func<IXlsxCell, Boolean> predicate = null)
    {
        var cellsUsed = CellsUsedInternal(options, r => r.LastCell(), predicate).ToList();

        if (cellsUsed.Count == 0)
            return null;

        var lastRow = cellsUsed.Max(c => c.Address.RowNumber);
        var lastColumn = cellsUsed.Max(c => c.Address.ColumnNumber);

        if (lastRow > RangeAddress.LastAddress.RowNumber)
            lastRow = RangeAddress.LastAddress.RowNumber;

        if (lastColumn > RangeAddress.LastAddress.ColumnNumber)
            lastColumn = RangeAddress.LastAddress.ColumnNumber;

        return Worksheet.Cell(lastRow, lastColumn);
    }

    internal XlsxCell Cell(Int32 row, Int32 column)
    {
        return Cell(new Point(row, column));
    }

    internal XlsxCell Cell(Point point)
    {
        return Cell(new XlsxAddress(Worksheet, point.Row, point.Column, false, false));
    }

    public virtual XlsxCell Cell(String cellAddressInRange)
    {
        if (XlsxHelper.IsValidA1Address(cellAddressInRange))
            return Cell(XlsxAddress.Create(Worksheet, cellAddressInRange));

        if (Worksheet.DefinedNames.TryGetValue(cellAddressInRange, out IXlsxDefinedName definedName))
            return definedName.Ranges.First().FirstCell().CastTo<XlsxCell>();

        return null;
    }

    public XlsxCell Cell(Int32 row, String column)
    {
        return Cell(new XlsxAddress(Worksheet, row, column, false, false));
    }

    public XlsxCell Cell(IXlsxAddress cellAddressInRange)
    {
        return Cell(cellAddressInRange.RowNumber, cellAddressInRange.ColumnNumber);
    }

    public XlsxCell Cell(in XlsxAddress cellAddressInRange)
    {
        Int32 absRow = cellAddressInRange.RowNumber + RangeAddress.FirstAddress.RowNumber - 1;
        Int32 absColumn = cellAddressInRange.ColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1;

        if (absRow <= 0 || absRow > XlsxHelper.MaxRowNumber)
        {
            throw new ArgumentOutOfRangeException(
                nameof(cellAddressInRange),
                String.Format("Row number must be between 1 and {0}", XlsxHelper.MaxRowNumber)
            );
        }

        if (absColumn <= 0 || absColumn > XlsxHelper.MaxColumnNumber)
        {
            throw new ArgumentOutOfRangeException(
                nameof(cellAddressInRange),
                String.Format("Column number must be between 1 and {0}", XlsxHelper.MaxColumnNumber)
            );
        }

        var cell = Worksheet.Internals.CellsCollection.GetCell(new Point(absRow, absColumn));
        return cell;
    }

    public Int32 RowCount()
    {
        return RangeAddress.LastAddress.RowNumber - RangeAddress.FirstAddress.RowNumber + 1;
    }

    public Int32 RowCount(XlsxCellsUsedOptions cellsUsedOptions)
    {
        var lcu = LastCellUsed(cellsUsedOptions);
        if (lcu == null) return 0;

        var fcu = FirstCellUsed(cellsUsedOptions);
        if (fcu == null) return 0;

        return lcu.Address.RowNumber - fcu.Address.RowNumber + 1;
    }

    public Int32 RowNumber()
    {
        return RangeAddress.FirstAddress.RowNumber;
    }

    public Int32 ColumnCount()
    {
        return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
    }

    public Int32 ColumnCount(XlsxCellsUsedOptions cellsUsedOptions)
    {
        var lcu = LastCellUsed(cellsUsedOptions);
        if (lcu == null) return 0;

        var fcu = FirstCellUsed(cellsUsedOptions);
        if (fcu == null) return 0;

        return lcu.Address.ColumnNumber - fcu.Address.ColumnNumber + 1;
    }

    public Int32 ColumnNumber()
    {
        return RangeAddress.FirstAddress.ColumnNumber;
    }

    public String ColumnLetter()
    {
        return RangeAddress.FirstAddress.ColumnLetter;
    }

    public virtual XlsxRange Range(String rangeAddressStr)
    {
        var rangeAddress = new XlsxRangeAddress(Worksheet, rangeAddressStr);
        return Range(rangeAddress);
    }

    internal abstract void WorksheetRangeShiftedColumns(XlsxRange range, int columnsShifted);

    internal abstract void WorksheetRangeShiftedRows(XlsxRange range, int rowsShifted);

    public abstract XlsxRangeType RangeType { get; }

    public XlsxRange Range(IXlsxCell firstCell, IXlsxCell lastCell)
    {
        var newFirstCellAddress = (XlsxAddress)firstCell.Address;
        var newLastCellAddress = (XlsxAddress)lastCell.Address;

        return GetRange(newFirstCellAddress, newLastCellAddress);
    }

    private XlsxRange GetRange(XlsxAddress newFirstCellAddress, XlsxAddress newLastCellAddress)
    {
        if (!this.Worksheet.Equals(newFirstCellAddress.Worksheet))
            throw new ArgumentException("The address refers to a different worksheet.", nameof(newFirstCellAddress));

        if (!this.Worksheet.Equals(newLastCellAddress.Worksheet))
            throw new ArgumentException("The address refers to a different worksheet.", nameof(newLastCellAddress));

        var newRangeAddress = new XlsxRangeAddress(newFirstCellAddress, newLastCellAddress);
        if (
            newFirstCellAddress.RowNumber < RangeAddress.FirstAddress.RowNumber
            || newFirstCellAddress.RowNumber > RangeAddress.LastAddress.RowNumber
            || newLastCellAddress.RowNumber > RangeAddress.LastAddress.RowNumber
            || newFirstCellAddress.ColumnNumber < RangeAddress.FirstAddress.ColumnNumber
            || newFirstCellAddress.ColumnNumber > RangeAddress.LastAddress.ColumnNumber
            || newLastCellAddress.ColumnNumber > RangeAddress.LastAddress.ColumnNumber
        )
        {
            throw new ArgumentOutOfRangeException(String.Format(
                "The cells {0} and {1} are outside the range '{2}'.",
                newFirstCellAddress,
                newLastCellAddress,
                ToString()));
        }

        if (newFirstCellAddress.Worksheet != null)
            return newFirstCellAddress.Worksheet.GetOrCreateRange(newRangeAddress);
        else if (Worksheet != null)
            return Worksheet.GetOrCreateRange(newRangeAddress);
        else
            return new XlsxRange(newRangeAddress);
    }

    public XlsxRange Range(String firstCellAddress, String lastCellAddress)
    {
        var rangeAddress = new XlsxRangeAddress(XlsxAddress.Create(Worksheet, firstCellAddress),
                                              XlsxAddress.Create(Worksheet, lastCellAddress));
        return Range(rangeAddress);
    }

    internal XlsxRange Range(Area area)
    {
        return Range(area.TopRow, area.LeftColumn, area.BottomRow, area.RightColumn);
    }

    internal XlsxRange Range(Int32 firstCellRow, Int32 firstCellColumn, Int32 lastCellRow, Int32 lastCellColumn)
    {
        var rangeAddress = new XlsxRangeAddress
        (
            new XlsxAddress
            (
                Worksheet,
                firstCellRow + RangeAddress.FirstAddress.RowNumber - 1,
                firstCellColumn + RangeAddress.FirstAddress.ColumnNumber - 1,
                fixedRow: false,
                fixedColumn: false
            ),
            new XlsxAddress
            (
                Worksheet,
                lastCellRow + RangeAddress.FirstAddress.RowNumber - 1,
                lastCellColumn + RangeAddress.FirstAddress.ColumnNumber - 1,
                fixedRow: false,
                fixedColumn: false
            )
        );
        return Range(rangeAddress);
    }

    public XlsxRange Range(IXlsxAddress firstCellAddress, IXlsxAddress lastCellAddress)
    {
        var rangeAddress = new XlsxRangeAddress((XlsxAddress)firstCellAddress, (XlsxAddress)lastCellAddress);
        return Range(rangeAddress);
    }

    public XlsxRange Range(IXlsxRangeAddress rangeAddress)
    {
        var xlRangeAddress = (XlsxRangeAddress)rangeAddress;
        return Range(in xlRangeAddress);
    }

    internal XlsxRange Range(in XlsxRangeAddress rangeAddress)
    {
        var ws = rangeAddress.FirstAddress.Worksheet ??
                 rangeAddress.LastAddress.Worksheet ??
                 Worksheet;

        var newFirstCellAddress = new XlsxAddress(ws,
                             rangeAddress.FirstAddress.RowNumber,
                             rangeAddress.FirstAddress.ColumnNumber,
                             rangeAddress.FirstAddress.FixedRow,
                             rangeAddress.FirstAddress.FixedColumn);

        var newLastCellAddress = new XlsxAddress(ws,
                            rangeAddress.LastAddress.RowNumber,
                            rangeAddress.LastAddress.ColumnNumber,
                            rangeAddress.LastAddress.FixedRow,
                            rangeAddress.LastAddress.FixedColumn);

        return GetRange(newFirstCellAddress, newLastCellAddress);
    }

    public virtual XlsxRanges Ranges(String ranges)
    {
        var retVal = new XlsxRanges(Worksheet);
        var rangePairs = ranges.Split(',');
        foreach (string pair in rangePairs)
            retVal.Add(Range(pair.Trim()));
        return retVal;
    }

    public IXlsxRanges Ranges(params String[] ranges)
    {
        var retVal = new XlsxRanges(Worksheet);
        foreach (string pair in ranges)
            retVal.Add(Range(pair));
        return retVal;
    }

    protected String FixColumnAddress(String address)
    {
        if (Int32.TryParse(address, out Int32 rowNumber))
            return RangeAddress.FirstAddress.ColumnLetter + (rowNumber + RangeAddress.FirstAddress.RowNumber - 1).ToInvariantString();
        return address;
    }

    protected String FixRowAddress(String address)
    {
        if (Int32.TryParse(address, out Int32 columnNumber))
            return XlsxHelper.GetColumnLetterFromNumber(columnNumber + RangeAddress.FirstAddress.ColumnNumber - 1) + RangeAddress.FirstAddress.RowNumber.ToInvariantString();
        return address;
    }

    public IXlsxCells CellsUsed(XlsxCellsUsedOptions options)
    {
        var cells = new XlsxCells(Worksheet, true, options) { RangeAddress };
        return cells;
    }

    public IXlsxCells CellsUsed(Func<IXlsxCell, Boolean> predicate)
    {
        var cells = new XlsxCells(Worksheet, true, XlsxCellsUsedOptions.AllContents, predicate) { RangeAddress };
        return cells;
    }

    public IXlsxCells CellsUsed(XlsxCellsUsedOptions options, Func<IXlsxCell, Boolean> predicate)
    {
        return CellsUsedInternal(options, predicate);
    }

    internal XlsxCells CellsUsedInternal(XlsxCellsUsedOptions options, Func<XlsxCell, Boolean> predicate = null)
    {
        var cells = new XlsxCells(Worksheet, true, options, predicate) { RangeAddress };
        return cells;
    }

    public IXlsxRangeColumns InsertColumnsAfter(Int32 numberOfColumns)
    {
        return InsertColumnsAfter(numberOfColumns, true);
    }

    public IXlsxRangeColumns InsertColumnsAfter(Int32 numberOfColumns, Boolean expandRange)
    {
        var retVal = InsertColumnsAfter(false, numberOfColumns);
        // Adjust the range
        if (expandRange)
        {
            RangeAddress = new XlsxRangeAddress(
                new XlsxAddress(Worksheet,
                              RangeAddress.FirstAddress.RowNumber,
                              RangeAddress.FirstAddress.ColumnNumber,
                              RangeAddress.FirstAddress.FixedRow,
                              RangeAddress.FirstAddress.FixedColumn),
                new XlsxAddress(Worksheet,
                              RangeAddress.LastAddress.RowNumber,
                              RangeAddress.LastAddress.ColumnNumber + numberOfColumns,
                              RangeAddress.LastAddress.FixedRow,
                              RangeAddress.LastAddress.FixedColumn));
        }
        return retVal;
    }

    public IXlsxRangeColumns InsertColumnsAfter(Boolean onlyUsedCells, Int32 numberOfColumns, Boolean formatFromLeft = true)
    {
        return InsertColumnsAfterInternal(onlyUsedCells, numberOfColumns, formatFromLeft);
    }

    public void InsertColumnsAfterVoid(Boolean onlyUsedCells, Int32 numberOfColumns, Boolean formatFromLeft = true)
    {
        InsertColumnsAfterInternal(onlyUsedCells, numberOfColumns, formatFromLeft, nullReturn: true);
    }

    private IXlsxRangeColumns InsertColumnsAfterInternal(Boolean onlyUsedCells, Int32 numberOfColumns, Boolean formatFromLeft = true, Boolean nullReturn = false)
    {
        int columnCount = ColumnCount();
        int firstColumn = RangeAddress.FirstAddress.ColumnNumber + columnCount;
        if (firstColumn > XlsxHelper.MaxColumnNumber)
            firstColumn = XlsxHelper.MaxColumnNumber;
        int lastColumn = firstColumn + ColumnCount() - 1;
        if (lastColumn > XlsxHelper.MaxColumnNumber)
            lastColumn = XlsxHelper.MaxColumnNumber;

        int firstRow = RangeAddress.FirstAddress.RowNumber;
        int lastRow = firstRow + RowCount() - 1;
        if (lastRow > XlsxHelper.MaxRowNumber)
            lastRow = XlsxHelper.MaxRowNumber;

        var newRange = Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
        return newRange.InsertColumnsBeforeInternal(onlyUsedCells, numberOfColumns, formatFromLeft, nullReturn);
    }

    public IXlsxRangeColumns InsertColumnsBefore(Int32 numberOfColumns)
    {
        return InsertColumnsBefore(numberOfColumns, false);
    }

    public IXlsxRangeColumns InsertColumnsBefore(Int32 numberOfColumns, Boolean expandRange)
    {
        var retVal = InsertColumnsBefore(false, numberOfColumns);
        // Adjust the range
        if (expandRange)
        {
            RangeAddress = new XlsxRangeAddress(
                new XlsxAddress(Worksheet,
                              RangeAddress.FirstAddress.RowNumber,
                              RangeAddress.FirstAddress.ColumnNumber - numberOfColumns,
                              RangeAddress.FirstAddress.FixedRow,
                              RangeAddress.FirstAddress.FixedColumn),
                new XlsxAddress(Worksheet,
                              RangeAddress.LastAddress.RowNumber,
                              RangeAddress.LastAddress.ColumnNumber,
                              RangeAddress.LastAddress.FixedRow,
                              RangeAddress.LastAddress.FixedColumn));
        }
        return retVal;
    }

    public IXlsxRangeColumns InsertColumnsBefore(Boolean onlyUsedCells, Int32 numberOfColumns, Boolean formatFromLeft = true)
    {
        return InsertColumnsBeforeInternal(onlyUsedCells, numberOfColumns, formatFromLeft);
    }

    public void InsertColumnsBeforeVoid(Boolean onlyUsedCells, Int32 numberOfColumns, Boolean formatFromLeft = true)
    {
        InsertColumnsBeforeInternal(onlyUsedCells, numberOfColumns, formatFromLeft, nullReturn: true);
    }

    private IXlsxRangeColumns InsertColumnsBeforeInternal(Boolean onlyUsedCells, Int32 numberOfColumns, Boolean formatFromLeft = true, Boolean nullReturn = false)
    {
        if (numberOfColumns <= 0 || numberOfColumns > XlsxHelper.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(numberOfColumns),
                $"Number of columns to insert must be a positive number no more than {XlsxHelper.MaxColumnNumber}");

        foreach (XlsxWorksheet ws in Worksheet.Workbook.WorksheetsInternal)
        {
            foreach (XlsxCell cell in ws.Internals.CellsCollection.GetCells(c => c.Formula is not null))
                cell.ShiftFormulaColumns(AsRange(), numberOfColumns);
        }

        // Inserting and shifting of whole columns is rather inconsistent across the codebase. In some places, the columns collection
        // is shifted before this method is called and thus the we can't shift column properties again. In others, the code relies on
        // shifting in this method.
        if (!onlyUsedCells)
        {
            int lastColumn = Worksheet.Internals.CellsCollection.MaxColumnUsed;
            if (lastColumn > 0)
            {
                int firstColumn = RangeAddress.FirstAddress.ColumnNumber;
                for (int co = lastColumn; co >= firstColumn; co--)
                {
                    int newColumn = co + numberOfColumns;
                    if (this.IsEntireColumn())
                    {
                        Worksheet.Column(newColumn).Width = Worksheet.Column(co).Width;
                    }
                }
            }
        }

        var insertedRange = new Area(
            Point.FromAddress(RangeAddress.FirstAddress),
            new Point(RangeAddress.LastAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber + numberOfColumns - 1));

        Worksheet.Internals.CellsCollection.InsertAreaAndShiftRight(insertedRange);

        Int32 firstRowReturn = RangeAddress.FirstAddress.RowNumber;
        Int32 lastRowReturn = RangeAddress.LastAddress.RowNumber;
        Int32 firstColumnReturn = RangeAddress.FirstAddress.ColumnNumber;
        Int32 lastColumnReturn = RangeAddress.FirstAddress.ColumnNumber + numberOfColumns - 1;

        Worksheet.NotifyRangeShiftedColumns(AsRange(), numberOfColumns);

        var rangeToReturn = Worksheet.Range(firstRowReturn, firstColumnReturn, lastRowReturn, lastColumnReturn);

        // We deliberately ignore conditional formats and data validation here. Their shifting is handled elsewhere
        var contentFlags = XlsxCellsUsedOptions.All
            & ~XlsxCellsUsedOptions.ConditionalFormats
            & ~XlsxCellsUsedOptions.DataValidation;

        if (formatFromLeft && rangeToReturn.RangeAddress.FirstAddress.ColumnNumber > 1)
        {
            var firstColumnUsed = rangeToReturn.FirstColumn();
            var model = firstColumnUsed.ColumnLeft();
            var modelFirstRow = (model as IXlsxRangeBase).FirstCellUsed(contentFlags);
            var modelLastRow = (model as IXlsxRangeBase).LastCellUsed(contentFlags);
            if (modelLastRow != null)
            {
                Int32 firstRoReturned = modelFirstRow.Address.RowNumber
                                        - model.RangeAddress.FirstAddress.RowNumber + 1;
                Int32 lastRoReturned = modelLastRow.Address.RowNumber
                                       - model.RangeAddress.FirstAddress.RowNumber + 1;
                for (Int32 ro = firstRoReturned; ro <= lastRoReturned; ro++)
                {
                    rangeToReturn.Row(ro).Style = model.Cell(ro).Style;
                }
            }
        }
        else
        {
            var lastRoUsed = rangeToReturn.LastRowUsed(contentFlags);
            if (lastRoUsed != null)
            {
                Int32 lastRoReturned = lastRoUsed.RowNumber();
                for (Int32 ro = 1; ro <= lastRoReturned; ro++)
                {
                    var styleToUse =
                        Worksheet.Internals.RowsCollection.TryGetValue(ro, out XlsxRow row)
                            ? row.Style
                            : Worksheet.Style;

                    rangeToReturn.Row(ro).Style = styleToUse;
                }
            }
        }

        if (nullReturn)
            return null;

        return rangeToReturn.Columns();
    }

    public IXlsxRangeRows InsertRowsBelow(Int32 numberOfRows)
    {
        return InsertRowsBelow(numberOfRows, true);
    }

    public IXlsxRangeRows InsertRowsBelow(Int32 numberOfRows, Boolean expandRange)
    {
        var retVal = InsertRowsBelow(false, numberOfRows);
        // Adjust the range
        if (expandRange)
        {
            RangeAddress = new XlsxRangeAddress(
                new XlsxAddress(Worksheet,
                              RangeAddress.FirstAddress.RowNumber,
                              RangeAddress.FirstAddress.ColumnNumber,
                              RangeAddress.FirstAddress.FixedRow,
                              RangeAddress.FirstAddress.FixedColumn),
                new XlsxAddress(Worksheet,
                              RangeAddress.LastAddress.RowNumber + numberOfRows,
                              RangeAddress.LastAddress.ColumnNumber,
                              RangeAddress.LastAddress.FixedRow,
                              RangeAddress.LastAddress.FixedColumn));
        }
        return retVal;
    }

    public IXlsxRangeRows InsertRowsBelow(Boolean onlyUsedCells, Int32 numberOfRows, Boolean formatFromAbove = true)
    {
        return InsertRowsBelowInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn: false);
    }

    public void InsertRowsBelowVoid(Boolean onlyUsedCells, Int32 numberOfRows, Boolean formatFromAbove = true)
    {
        InsertRowsBelowInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn: true);
    }

    private IXlsxRangeRows InsertRowsBelowInternal(Boolean onlyUsedCells, Int32 numberOfRows, Boolean formatFromAbove, Boolean nullReturn)
    {
        int rowCount = RowCount();
        int firstRow = RangeAddress.FirstAddress.RowNumber + rowCount;
        if (firstRow > XlsxHelper.MaxRowNumber)
            firstRow = XlsxHelper.MaxRowNumber;
        int lastRow = firstRow + RowCount() - 1;
        if (lastRow > XlsxHelper.MaxRowNumber)
            lastRow = XlsxHelper.MaxRowNumber;

        int firstColumn = RangeAddress.FirstAddress.ColumnNumber;
        int lastColumn = firstColumn + ColumnCount() - 1;
        if (lastColumn > XlsxHelper.MaxColumnNumber)
            lastColumn = XlsxHelper.MaxColumnNumber;

        var newRange = Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
        return newRange.InsertRowsAboveInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn);
    }

    public IXlsxRangeRows InsertRowsAbove(Int32 numberOfRows)
    {
        return InsertRowsAbove(numberOfRows, false);
    }

    public IXlsxRangeRows InsertRowsAbove(Int32 numberOfRows, Boolean expandRange)
    {
        var retVal = InsertRowsAbove(false, numberOfRows);
        // Adjust the range
        if (expandRange)
        {
            RangeAddress = new XlsxRangeAddress(
                new XlsxAddress(Worksheet,
                              RangeAddress.FirstAddress.RowNumber - numberOfRows,
                              RangeAddress.FirstAddress.ColumnNumber,
                              RangeAddress.FirstAddress.FixedRow,
                              RangeAddress.FirstAddress.FixedColumn),
                new XlsxAddress(Worksheet,
                              RangeAddress.LastAddress.RowNumber,
                              RangeAddress.LastAddress.ColumnNumber,
                              RangeAddress.LastAddress.FixedRow,
                              RangeAddress.LastAddress.FixedColumn));
        }
        return retVal;
    }

    public void InsertRowsAboveVoid(Boolean onlyUsedCells, Int32 numberOfRows, Boolean formatFromAbove = true)
    {
        InsertRowsAboveInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn: true);
    }

    public IXlsxRangeRows InsertRowsAbove(Boolean onlyUsedCells, Int32 numberOfRows, Boolean formatFromAbove = true)
    {
        return InsertRowsAboveInternal(onlyUsedCells, numberOfRows, formatFromAbove, nullReturn: false);
    }

    private IXlsxRangeRows InsertRowsAboveInternal(Boolean onlyUsedCells, Int32 numberOfRows, Boolean formatFromAbove, Boolean nullReturn)
    {
        if (numberOfRows <= 0 || numberOfRows > XlsxHelper.MaxRowNumber)
            throw new ArgumentOutOfRangeException(nameof(numberOfRows),
                $"Number of rows to insert must be a positive number no more than {XlsxHelper.MaxRowNumber}");

        var asRange = AsRange();
        foreach (XlsxWorksheet ws in Worksheet.Workbook.WorksheetsInternal)
        {
            foreach (XlsxCell cell in ws.Internals.CellsCollection.GetCells(c => c.Formula is not null))
                cell.ShiftFormulaRows(asRange, numberOfRows);
        }

        if (!onlyUsedCells)
        {
            int lastRow = Worksheet.Internals.CellsCollection.MaxRowUsed;
            if (lastRow > 0)
            {
                int firstRow = RangeAddress.FirstAddress.RowNumber;
                for (var ro = lastRow; ro >= firstRow; ro--)
                {
                    var newRow = ro + numberOfRows;
                    if (this.IsEntireRow())
                    {
                        Worksheet.Row(newRow).Height = Worksheet.Row(ro).Height;
                    }
                }
            }
        }

        var insertedRange = new Area(
            Point.FromAddress(RangeAddress.FirstAddress),
            new Point(RangeAddress.FirstAddress.RowNumber + numberOfRows - 1, RangeAddress.LastAddress.ColumnNumber));
        Worksheet.Internals.CellsCollection.InsertAreaAndShiftDown(insertedRange);

        Int32 firstRowReturn = RangeAddress.FirstAddress.RowNumber;
        Int32 lastRowReturn = RangeAddress.FirstAddress.RowNumber + numberOfRows - 1;
        Int32 firstColumnReturn = RangeAddress.FirstAddress.ColumnNumber;
        Int32 lastColumnReturn = RangeAddress.LastAddress.ColumnNumber;

        Worksheet.NotifyRangeShiftedRows(AsRange(), numberOfRows);

        var rangeToReturn = Worksheet.Range(firstRowReturn, firstColumnReturn, lastRowReturn, lastColumnReturn);

        // We deliberately ignore conditional formats and data validation here. Their shifting is handled elsewhere
        var contentFlags = XlsxCellsUsedOptions.All
            & ~XlsxCellsUsedOptions.ConditionalFormats
            & ~XlsxCellsUsedOptions.DataValidation;

        if (formatFromAbove && rangeToReturn.RangeAddress.FirstAddress.RowNumber > 1)
        {
            var fr = rangeToReturn.FirstRow();
            var model = fr.RowAbove();
            var modelFirstColumn = (model as IXlsxRangeBase).FirstCellUsed(contentFlags);
            var modelLastColumn = (model as IXlsxRangeBase).LastCellUsed(contentFlags);
            if (modelFirstColumn != null && modelLastColumn != null)
            {
                Int32 firstCoReturned = modelFirstColumn.Address.ColumnNumber
                                        - model.RangeAddress.FirstAddress.ColumnNumber + 1;
                Int32 lastCoReturned = modelLastColumn.Address.ColumnNumber
                                        - model.RangeAddress.FirstAddress.ColumnNumber + 1;
                for (Int32 co = firstCoReturned; co <= lastCoReturned; co++)
                {
                    rangeToReturn.Column(co).Style = model.Cell(co).Style;
                }
            }
        }
        else
        {
            var lastCoUsed = rangeToReturn.LastColumnUsed(contentFlags);
            if (lastCoUsed != null)
            {
                Int32 lastCoReturned = lastCoUsed.ColumnNumber();
                for (Int32 co = 1; co <= lastCoReturned; co++)
                {
                    var styleToUse =
                        Worksheet.Internals.ColumnsCollection.TryGetValue(co, out XlsxColumn column)
                            ? column.Style
                            : Worksheet.Style;

                    rangeToReturn.Style = styleToUse;
                }
            }
        }

        // Skip calling .Rows() for performance reasons if required.
        if (nullReturn)
            return null;

        return rangeToReturn.Rows();
    }

    private void ClearMerged()
    {
        var mergeToDelete = Worksheet.Internals.MergedRanges.GetIntersectedRanges(RangeAddress).ToList();
        mergeToDelete.ForEach(m => Worksheet.Internals.MergedRanges.Remove(m));
    }

    public Boolean Contains(IXlsxCell cell)
    {
        return Contains((XlsxAddress)cell.Address);
    }

    public bool Contains(XlsxAddress first, XlsxAddress last)
    {
        return Contains(first) && Contains(last);
    }

    public bool Contains(XlsxAddress address)
    {
        return RangeAddress.Contains(in address);
    }

    public void Delete(XlsxShiftDeletedCells shiftDeleteCells)
    {
        int numberOfRows = RowCount();
        int numberOfColumns = ColumnCount();

        if (!RangeAddress.IsValid) return;

        Worksheet.SparklineGroups.Remove(this);

        IXlsxRange shiftedRangeFormula = Worksheet.Range(
            RangeAddress.FirstAddress.RowNumber,
            RangeAddress.FirstAddress.ColumnNumber,
            RangeAddress.LastAddress.RowNumber,
            RangeAddress.LastAddress.ColumnNumber);

        // Shift formulas first
        foreach (var cell in Worksheet
            .Workbook
            .Worksheets
            .Cast<XlsxWorksheet>()
            .SelectMany(ws => ws
                .Internals
                .CellsCollection
                .GetCells(c => c.HasFormula)))
        {
            if (shiftDeleteCells == XlsxShiftDeletedCells.ShiftCellsUp)
                cell.ShiftFormulaRows((XlsxRange)shiftedRangeFormula, numberOfRows * -1);
            else
                cell.ShiftFormulaColumns((XlsxRange)shiftedRangeFormula, numberOfColumns * -1);
        }

        // Range to shift...
        Int32 columnModifier = 0;
        Int32 rowModifier = 0;
        var range = Area.FromRangeAddress(RangeAddress);
        switch (shiftDeleteCells)
        {
            case XlsxShiftDeletedCells.ShiftCellsLeft:
                Worksheet.Internals.CellsCollection.DeleteAreaAndShiftLeft(range);
                columnModifier = ColumnCount();
                break;

            case XlsxShiftDeletedCells.ShiftCellsUp:
                Worksheet.Internals.CellsCollection.DeleteAreaAndShiftUp(range);
                rowModifier = RowCount();
                break;
        }

        var mergesToRemove = Worksheet.Internals.MergedRanges.Where<XlsxRange>(Contains).ToList();
        mergesToRemove.ForEach(r => Worksheet.Internals.MergedRanges.Remove(r));

        var shiftedRange = AsRange();
        if (shiftDeleteCells == XlsxShiftDeletedCells.ShiftCellsUp)
            Worksheet.NotifyRangeShiftedRows(shiftedRange, rowModifier * -1);
        else
            Worksheet.NotifyRangeShiftedColumns(shiftedRange, columnModifier * -1);

        Worksheet.DeleteRange(RangeAddress);
    }

    public override string ToString()
    {
        return String.Concat(
            Worksheet.Name.EscapeSheetName(),
            '!',
            RangeAddress.FirstAddress,
            ':',
            RangeAddress.LastAddress);
    }

    protected IXlsxRangeAddress ShiftColumns(IXlsxRangeAddress thisRangeAddress, XlsxRange shiftedRange, int columnsShifted)
    {
        if (!thisRangeAddress.IsValid || !shiftedRange.RangeAddress.IsValid) return thisRangeAddress;

        bool allRowsAreCovered = thisRangeAddress.FirstAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber &&
                                 thisRangeAddress.LastAddress.RowNumber <= shiftedRange.RangeAddress.LastAddress.RowNumber;

        if (!allRowsAreCovered)
            return thisRangeAddress;

        bool shiftLeftBoundary = (columnsShifted > 0 && thisRangeAddress.FirstAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber) ||
                                 (columnsShifted < 0 && thisRangeAddress.FirstAddress.ColumnNumber > shiftedRange.RangeAddress.FirstAddress.ColumnNumber);

        bool shiftRightBoundary = thisRangeAddress.LastAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber;

        int newLeftBoundary = thisRangeAddress.FirstAddress.ColumnNumber;
        if (shiftLeftBoundary)
        {
            if (newLeftBoundary + columnsShifted > shiftedRange.RangeAddress.FirstAddress.ColumnNumber)
                newLeftBoundary += columnsShifted;
            else
                newLeftBoundary = shiftedRange.RangeAddress.FirstAddress.ColumnNumber;
        }

        int newRightBoundary = thisRangeAddress.LastAddress.ColumnNumber;
        if (shiftRightBoundary)
            newRightBoundary += columnsShifted;

        bool destroyedByShift = newRightBoundary < newLeftBoundary;

        var firstAddress = (XlsxAddress)thisRangeAddress.FirstAddress;
        var lastAddress = (XlsxAddress)thisRangeAddress.LastAddress;

        if (destroyedByShift)
        {
            firstAddress = Worksheet.InvalidAddress;
            lastAddress = Worksheet.InvalidAddress;
            Worksheet.DeleteRange(RangeAddress);
        }

        if (shiftLeftBoundary)
            firstAddress = new XlsxAddress(Worksheet,
                                         thisRangeAddress.FirstAddress.RowNumber,
                                         newLeftBoundary,
                                         thisRangeAddress.FirstAddress.FixedRow,
                                         thisRangeAddress.FirstAddress.FixedColumn);

        if (shiftRightBoundary)
            lastAddress = new XlsxAddress(Worksheet,
                                        thisRangeAddress.LastAddress.RowNumber,
                                        newRightBoundary,
                                        thisRangeAddress.LastAddress.FixedRow,
                                        thisRangeAddress.LastAddress.FixedColumn);

        return new XlsxRangeAddress(firstAddress, lastAddress);
    }

    protected IXlsxRangeAddress ShiftRows(IXlsxRangeAddress thisRangeAddress, XlsxRange shiftedRange, int rowsShifted)
    {
        if (!thisRangeAddress.IsValid || !shiftedRange.RangeAddress.IsValid) return thisRangeAddress;

        bool allColumnsAreCovered = thisRangeAddress.FirstAddress.ColumnNumber >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber &&
                                    thisRangeAddress.LastAddress.ColumnNumber <= shiftedRange.RangeAddress.LastAddress.ColumnNumber;

        if (!allColumnsAreCovered)
            return thisRangeAddress;

        bool shiftTopBoundary = (rowsShifted > 0 && thisRangeAddress.FirstAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber) ||
                                (rowsShifted < 0 && thisRangeAddress.FirstAddress.RowNumber > shiftedRange.RangeAddress.FirstAddress.RowNumber);

        bool shiftBottomBoundary = thisRangeAddress.LastAddress.RowNumber >= shiftedRange.RangeAddress.FirstAddress.RowNumber;

        int newTopBoundary = thisRangeAddress.FirstAddress.RowNumber;
        if (shiftTopBoundary)
        {
            if (newTopBoundary + rowsShifted > shiftedRange.RangeAddress.FirstAddress.RowNumber)
                newTopBoundary += rowsShifted;
            else
                newTopBoundary = shiftedRange.RangeAddress.FirstAddress.RowNumber;
        }

        int newBottomBoundary = thisRangeAddress.LastAddress.RowNumber;
        if (shiftBottomBoundary)
            newBottomBoundary += rowsShifted;

        bool destroyedByShift = newBottomBoundary < newTopBoundary;

        var firstAddress = (XlsxAddress)thisRangeAddress.FirstAddress;
        var lastAddress = (XlsxAddress)thisRangeAddress.LastAddress;

        if (destroyedByShift)
        {
            firstAddress = Worksheet.InvalidAddress;
            lastAddress = Worksheet.InvalidAddress;
            Worksheet.DeleteRange(RangeAddress);
        }

        if (shiftTopBoundary)
            firstAddress = new XlsxAddress(Worksheet,
                                         newTopBoundary,
                                         thisRangeAddress.FirstAddress.ColumnNumber,
                                         thisRangeAddress.FirstAddress.FixedRow,
                                         thisRangeAddress.FirstAddress.FixedColumn);

        if (shiftBottomBoundary)
            lastAddress = new XlsxAddress(Worksheet,
                                        newBottomBoundary,
                                        thisRangeAddress.LastAddress.ColumnNumber,
                                        thisRangeAddress.LastAddress.FixedRow,
                                        thisRangeAddress.LastAddress.FixedColumn);

        return new XlsxRangeAddress(firstAddress, lastAddress);
    }

    public IXlsxRange RangeUsed()
    {
        return RangeUsed(XlsxCellsUsedOptions.AllContents);
    }

    public IXlsxRange RangeUsed(XlsxCellsUsedOptions options)
    {
        var firstCell = (this as IXlsxRangeBase).FirstCellUsed(options);
        if (firstCell == null)
            return null;
        var lastCell = (this as IXlsxRangeBase).LastCellUsed(options);
        return Worksheet.Range(firstCell, lastCell);
    }

    public virtual void CopyTo(IXlsxRangeBase target)
    {
        CopyTo((XlsxCell)target.FirstCell());
    }

    internal void CopyTo(XlsxCell target)
    {
        target.CopyFrom(this);
    }

    //public IXlsxChart CreateChart(Int32 firstRow, Int32 firstColumn, Int32 lastRow, Int32 lastColumn)
    //{
    //    IXlsxChart chart = new XlsxChartWorksheet;
    //    chart.FirstRow = firstRow;
    //    chart.LastRow = lastRow;
    //    chart.LastColumn = lastColumn;
    //    chart.FirstColumn = firstColumn;
    //    Worksheet.Charts.Add(chart);
    //    return chart;
    //}

    IXlsxPivotTable IXlsxRangeBase.CreatePivotTable(IXlsxCell targetCell, String name)
    {
        return CreatePivotTable(targetCell, name);
    }

    public XlsxPivotTable CreatePivotTable(IXlsxCell targetCell, String name)
    {
        return (XlsxPivotTable)targetCell.Worksheet.PivotTables.Add(name, targetCell, AsRange());
    }

    public virtual IXlsxAutoFilter SetAutoFilter()
    {
        return SetAutoFilter(true);
    }

    public IXlsxAutoFilter SetAutoFilter(Boolean value)
    {
        if (value)
            return Worksheet.AutoFilter.Set(this);
        else
            return Worksheet.AutoFilter.Clear();
    }

    #region Sort

    public IXlsxSortElements SortRows
    {
        get { return _sortRows ??= []; }
    }

    public IXlsxSortElements SortColumns
    {
        get { return _sortColumns ??= []; }
    }

    private String DefaultSortString()
    {
        var sb = new StringBuilder();
        Int32 maxColumn = ColumnCount();
        if (maxColumn == XlsxHelper.MaxColumnNumber)
            maxColumn = (this as IXlsxRangeBase).LastCellUsed(XlsxCellsUsedOptions.All).Address.ColumnNumber;
        for (int i = 1; i <= maxColumn; i++)
        {
            if (sb.Length > 0)
                sb.Append(',');

            sb.Append(i);
        }

        return sb.ToString();
    }

    public IXlsxRangeBase Sort()
    {
        if (!SortColumns.Any())
        {
            return Sort(DefaultSortString());
        }

        SortRangeRows();
        return this;
    }

    public IXlsxRangeBase Sort(String columnsToSortBy, XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
    {
        SortColumns.Clear();
        if (String.IsNullOrWhiteSpace(columnsToSortBy))
        {
            columnsToSortBy = DefaultSortString();
        }

        SortColumns.CastTo<XlsxSortElements>().AddRange(ParseSortOrder(columnsToSortBy, sortOrder, matchCase, ignoreBlanks));

        SortRangeRows();
        return this;
    }

    public IXlsxRangeBase Sort(Int32 columnToSortBy, XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
    {
        return Sort(columnToSortBy.ToString(), sortOrder, matchCase, ignoreBlanks);
    }

    public IXlsxRangeBase SortLeftToRight(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
    {
        SortRows.Clear();
        Int32 maxColumn = ColumnCount();
        if (maxColumn == XlsxHelper.MaxColumnNumber)
            maxColumn = (this as IXlsxRangeBase).LastCellUsed(XlsxCellsUsedOptions.All).Address.ColumnNumber;

        for (int i = 1; i <= maxColumn; i++)
        {
            SortRows.Add(i, sortOrder, ignoreBlanks, matchCase);
        }

        SortRangeColumns();
        return this;
    }

    private void SortRangeRows()
    {
        var sortRange = SheetRange;
        var cellsCollection = Worksheet.Internals.CellsCollection;
        if (sortRange.IsEntireColumn())
        {
            // If we're dealing with the entire column, we're not interested in the unused cells
            var lastRowUsed = cellsCollection.LastRowUsed(Area.Full, XlsxCellsUsedOptions.Contents);
            sortRange = new Area(sortRange.FirstPoint, new Point(lastRowUsed, sortRange.RightColumn));
        }

        var comparer = new XlsxRangeRowsSortComparer(Worksheet, sortRange, SortColumns);
        var rows = new int[sortRange.Height];
        for (var i = 0; i < sortRange.Height; ++i)
            rows[i] = i + sortRange.TopRow;

        Array.Sort(rows, comparer);

        cellsCollection.RemapRows(rows, sortRange);
    }

    private void SortRangeColumns()
    {
        var sortRange = SheetRange;
        var cellsCollection = Worksheet.Internals.CellsCollection;
        if (sortRange.IsEntireRow())
        {
            // If we're dealing with the entire row, we're not interested in the unused cells
            var lastColumnCell = cellsCollection.LastColumnUsed(Area.Full, XlsxCellsUsedOptions.Contents);
            sortRange = new Area(sortRange.FirstPoint, new Point(sortRange.BottomRow, lastColumnCell));
        }

        var comparer = new XlsxRangeColumnsSortComparer(Worksheet, sortRange, SortRows);
        var columns = new int[sortRange.Width];
        for (var i = 0; i < sortRange.Width; ++i)
            columns[i] = i + sortRange.LeftColumn;

        Array.Sort(columns, comparer);

        cellsCollection.RemapColumns(columns, sortRange);
    }

    private static IEnumerable<XlsxSortElement> ParseSortOrder(string columnsToSortBy, XlsxSortOrder defaultSortOrder, bool matchCase, bool ignoreBlanks)
    {
        foreach (var sortColumn in columnsToSortBy.Split(',').Select(coPair => coPair.Trim()))
        {
            var sortOrder = defaultSortOrder;

            String columnName;
            if (sortColumn.Contains(' '))
            {
                var pair = sortColumn.Split(' ');
                columnName = pair[0];
                sortOrder = pair[1].Equals("ASC", StringComparison.OrdinalIgnoreCase) ? XlsxSortOrder.Ascending : XlsxSortOrder.Descending;
            }
            else
            {
                columnName = sortColumn;
            }

            if (!Int32.TryParse(columnName, out Int32 columnNumber))
                columnNumber = XlsxHelper.GetColumnNumberFromLetter(columnName);

            yield return new XlsxSortElement(
                columnNumber,
                sortOrder,
                ignoreBlanks,
                matchCase);
        }
    }

    #endregion Sort

    public XlsxRangeColumn ColumnQuick(Int32 column)
    {
        var firstCellAddress = new XlsxAddress(Worksheet,
                                             RangeAddress.FirstAddress.RowNumber,
                                             RangeAddress.FirstAddress.ColumnNumber + column - 1,
                                             false,
                                             false);
        var lastCellAddress = new XlsxAddress(Worksheet,
                                            RangeAddress.LastAddress.RowNumber,
                                            RangeAddress.FirstAddress.ColumnNumber + column - 1,
                                            false,
                                            false);
        return Worksheet.RangeColumn(new XlsxRangeAddress(firstCellAddress, lastCellAddress));
    }

    public IXlsxConditionalFormat AddConditionalFormat()
    {
        var cf = new XlsxConditionalFormat(Worksheet, SheetRange.ToAreaList());
        Worksheet.ConditionalFormats.Add(cf);
        return cf;
    }

    public void Select()
    {
        Worksheet.SelectedRanges.Add(AsRange());
    }

    public IXlsxRangeBase Grow()
    {
        return Grow(1);
    }

    public IXlsxRangeBase Grow(int growCount)
    {
        var firstRow = Math.Max(1, this.RangeAddress.FirstAddress.RowNumber - growCount);
        var firstColumn = Math.Max(1, this.RangeAddress.FirstAddress.ColumnNumber - growCount);

        var lastRow = Math.Min(XlsxHelper.MaxRowNumber, this.RangeAddress.LastAddress.RowNumber + growCount);
        var lastColumn = Math.Min(XlsxHelper.MaxColumnNumber, this.RangeAddress.LastAddress.ColumnNumber + growCount);

        return this.Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
    }

    public IXlsxRangeBase Shrink()
    {
        return Shrink(1);
    }

    public IXlsxRangeBase Shrink(int shrinkCount)
    {
        var firstRow = this.RangeAddress.FirstAddress.RowNumber + shrinkCount;
        var firstColumn = this.RangeAddress.FirstAddress.ColumnNumber + shrinkCount;

        var lastRow = this.RangeAddress.LastAddress.RowNumber - shrinkCount;
        var lastColumn = this.RangeAddress.LastAddress.ColumnNumber - shrinkCount;

        if (firstRow > lastRow || firstColumn > lastColumn)
            return null;

        return this.Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
    }

    public IXlsxRangeAddress Intersection(IXlsxRangeBase otherRange, Func<IXlsxCell, Boolean> thisRangePredicate = null, Func<IXlsxCell, Boolean> otherRangePredicate = null)
    {
        if (otherRange == null)
            return null;

        if (!this.Worksheet.Equals(otherRange.Worksheet))
            return null;

        if (thisRangePredicate == null && otherRangePredicate == null)
        {
            // Special case, no predicates. We can optimise this a bit then.
            return this.RangeAddress.Intersection(otherRange.RangeAddress);
        }
        else
        {
            thisRangePredicate ??= (c => true);
            otherRangePredicate ??= (c => true);

            var intersectionCells = this.Cells(c => thisRangePredicate(c) && otherRange.Cells(otherRangePredicate).Contains(c));

            if (!intersectionCells.Any())
                return null;

            var firstRow = intersectionCells.Min(c => c.Address.RowNumber);
            var firstColumn = intersectionCells.Min(c => c.Address.ColumnNumber);

            var lastRow = intersectionCells.Max(c => c.Address.RowNumber);
            var lastColumn = intersectionCells.Max(c => c.Address.ColumnNumber);

            return new XlsxRangeAddress
            (
                new XlsxAddress(this.Worksheet, firstRow, firstColumn, fixedRow: false, fixedColumn: false),
                new XlsxAddress(this.Worksheet, lastRow, lastColumn, fixedRow: false, fixedColumn: false)
            );
        }
    }

    public IXlsxCells SurroundingCells(Func<IXlsxCell, Boolean> predicate = null)
    {
        var cells = new XlsxCells(Worksheet, false, XlsxCellsUsedOptions.AllContents, predicate);
        this.Grow().Cells(c => !this.Contains(c)).ForEach(c => cells.Add(c as XlsxCell));
        return cells;
    }

    public IXlsxCells Union(IXlsxRangeBase otherRange, Func<IXlsxCell, Boolean> thisRangePredicate = null, Func<IXlsxCell, Boolean> otherRangePredicate = null)
    {
        if (otherRange == null)
            return this.Cells(thisRangePredicate);

        var cells = new XlsxCells(Worksheet, false, XlsxCellsUsedOptions.AllContents);
        if (!this.Worksheet.Equals(otherRange.Worksheet))
            return cells;

        thisRangePredicate ??= c => true;
        otherRangePredicate ??= c => true;

        this.Cells(thisRangePredicate).Concat(otherRange.Cells(otherRangePredicate)).Distinct().ForEach(c => cells.Add(c as XlsxCell));
        return cells;
    }

    public IXlsxCells Difference(IXlsxRangeBase otherRange, Func<IXlsxCell, Boolean> thisRangePredicate = null, Func<IXlsxCell, Boolean> otherRangePredicate = null)
    {
        if (otherRange == null)
            return this.Cells(thisRangePredicate);

        var cells = new XlsxCells(Worksheet, false, XlsxCellsUsedOptions.AllContents);
        if (!this.Worksheet.Equals(otherRange.Worksheet))
            return cells;

        thisRangePredicate ??= c => true;
        otherRangePredicate ??= c => true;

        this.Cells(c => thisRangePredicate(c) && !otherRange.Cells(otherRangePredicate).Contains(c)).ForEach(c => cells.Add(c as XlsxCell));
        return cells;
    }

    private IEnumerable<IXlsxCell> CellsUsedInternal(XlsxCellsUsedOptions options, Func<IXlsxRange, IXlsxCell> selector, Func<IXlsxCell, bool> predicate)
    {
        predicate ??= (t => true);

        // To avoid unnecessary initialization of thousands cells to not hang on very large CFs, DVs or merged ranges.
        var opt = options
                  & ~XlsxCellsUsedOptions.ConditionalFormats
                  & ~XlsxCellsUsedOptions.DataValidation
                  & ~XlsxCellsUsedOptions.MergedRanges;

        // If opt == 0 then we're basically back at unconstrained, so just set back the original options
        if (opt == XlsxCellsUsedOptions.NoConstraints)
            opt = options;

        IEnumerable<IXlsxCell> cellsUsed = CellsUsed(opt, predicate);

        if (options.HasFlag(XlsxCellsUsedOptions.ConditionalFormats))
        {
            var area = SheetRange;
            cellsUsed = cellsUsed.Union(
                Worksheet.ConditionalFormats
                    .SelectMany<XlsxConditionalFormat, Area>(cf => cf.Areas.IntersectingWith(area))
                    .Select(cfArea => Worksheet.Range(cfArea))
                    .Select(selector)
                    .Where(predicate)
            );
        }
        if (options.HasFlag(XlsxCellsUsedOptions.DataValidation))
        {
            cellsUsed = cellsUsed.Union(
                Worksheet.DataValidations
                    .GetAllInRange(RangeAddress)
                    .SelectMany(dv => dv.Ranges)
                    .Select(selector)
                    .Where(predicate)
            );
        }
        if (options.HasFlag(XlsxCellsUsedOptions.MergedRanges))
        {
            cellsUsed = cellsUsed.Union(
                Worksheet.MergedRanges.GetIntersectedRanges(RangeAddress)
                    .Select(selector)
                    .Where(predicate)
            );
        }

        return cellsUsed;
    }
}
