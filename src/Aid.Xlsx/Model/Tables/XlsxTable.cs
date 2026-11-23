#nullable disable

using System.Collections;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Diagnostics;
using System.Dynamic;
using System.Globalization;
using System.Text;
using Array = System.Array;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.AutoFilters;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Ranges.Sort;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Tables;

[DebuggerDisplay("{Name}")]
internal class XlsxTable : XlsxRange, IXlsxTable
{
    private string _name;
    internal bool _showTotalsRow;
    internal HashSet<string> _uniqueNames;

    /// <summary>
    /// The direct constructor should only be used in <see cref="XlsxWorksheet.RangeFactory"/>.
    /// </summary>
    public XlsxTable(XlsxRangeAddress tableArea, IXlsxStyle defaultStyle)
        : base(tableArea)
    {
        InitializeValues(false);
    }
    public override XlsxRangeType RangeType
    {
        get { return XlsxRangeType.Table; }
    }

    private IXlsxRangeAddress _lastRangeAddress;
    private Dictionary<string, XlsxTableField> _fieldNames = CreateFieldNames();

    internal Dictionary<string, XlsxTableField> FieldNames
    {
        get
        {
            if (_fieldNames != null && _lastRangeAddress != null && _lastRangeAddress.Equals(RangeAddress))
                return _fieldNames;

            _lastRangeAddress = RangeAddress;

            RescanFieldNames();

            return _fieldNames;
        }
    }

    /// <summary>
    /// Area of the range, including headings and totals, if table has them.
    /// </summary>
    internal Area Area => Area.FromRangeAddress(RangeAddress);

    private void RescanFieldNames()
    {
        if (ShowHeaderRow)
        {
            Dictionary<string, XlsxTableField> oldFieldNames = _fieldNames ?? CreateFieldNames();
            _fieldNames = CreateFieldNames();
            XlsxRangeRow headersRow = HeadersRow(false);
            int cellPos = 0;
            foreach (XlsxCell cell in headersRow.Cells().Cast<XlsxCell>())
            {
                XlsxCellValue cellValue = cell.CachedValue;
                var name = cellValue.ToString(CultureInfo.CurrentCulture);

                if (oldFieldNames.TryGetValue(name, out XlsxTableField tableField))// && tableField.Column.ColumnNumber() == cell.Address.ColumnNumber)
                {
                    tableField.Index = cellPos;
                    _fieldNames.Add(name, tableField);
                    cellPos++;
                    continue;
                }

                // Be careful here. Fields names may actually be whitespace, but not empty
                if (string.IsNullOrEmpty(name))
                {
                    name = GetUniqueName("Column", cellPos + 1, true);
                }
                if (_fieldNames.ContainsKey(name))
                    throw new ArgumentException("The header row contains more than one field name '" + name + "'.");

                _fieldNames.Add(name, new XlsxTableField(this, name) { Index = cellPos++ });

                // Field names are the source of the truth that is projected
                // to the cells and field names can be only text. Fix the cell,
                // so cell fulfills its job of being dependent on the field name.
                if (!cellValue.Equals(name))
                {
                    cell.SetValue(name, false, false);
                }
            }
        }
        else
        {
            int colCount = ColumnCount();
            for (int i = 1; i <= colCount; i++)
            {
                if (_fieldNames.Values.All(f => f.Index != i - 1))
                {
                    var name = "Column" + i;

                    _fieldNames.Add(name, new XlsxTableField(this, name) { Index = i - 1 });
                }
            }
        }
    }

    internal XlsxTableField AddField(string fieldName)
    {
        var field = new XlsxTableField(this, fieldName) { Index = _fieldNames.Count };
        _fieldNames.Add(fieldName, field);
        return field;
    }

    internal void RenameField(string oldName, string newName)
    {
        if (!_fieldNames.TryGetValue(oldName, out XlsxTableField field))
            throw new ArgumentException("The field does not exist in this table", nameof(oldName));

        _fieldNames.Remove(oldName);
        _fieldNames.Add(newName, field);
    }

    internal string RelId { get; set; }

    public IXlsxTableRange DataRange
    {
        get
        {
            XlsxRange range;

            var firstDataRowNumber = 1;
            var lastDataRowNumber = RowCount();

            if (_showHeaderRow)
                firstDataRowNumber++;

            if (_showTotalsRow)
                lastDataRowNumber--;

            if (firstDataRowNumber > lastDataRowNumber)
                return null;

            range = Range(firstDataRowNumber, 1, lastDataRowNumber, ColumnCount());

            return new XlsxTableRange(range, this);
        }
    }

    private XlsxAutoFilter _autoFilter;

    public XlsxAutoFilter AutoFilter
    {
        get
        {
            _autoFilter ??= new XlsxAutoFilter();

            _autoFilter.Range = ShowTotalsRow ? Range(1, 1, RowCount() - 1, ColumnCount()) : AsRange();
            return _autoFilter;
        }
    }

    public override IXlsxAutoFilter SetAutoFilter()
    {
        return AutoFilter;
    }

    protected override void OnRangeAddressChanged(XlsxRangeAddress oldAddress, XlsxRangeAddress newAddress)
    {
        //Do nothing for table
    }

    #region IXlsxTable Members

    public bool EmphasizeFirstColumn { get; set; }
    public bool EmphasizeLastColumn { get; set; }
    public bool ShowRowStripes { get; set; }
    public bool ShowColumnStripes { get; set; }

    private bool _showAutoFilter;

    public bool ShowAutoFilter
    {
        get { return _showHeaderRow && _showAutoFilter; }
        set { _showAutoFilter = value; }
    }

    public XlsxTableTheme Theme { get; set; }

    public string Name
    {
        get { return _name; }
        set
        {
            if (_name == value) return;

            // Validation rules for table names
            var oldname = _name ?? string.Empty;
            IEnumerable<string> tableNames = Worksheet.Tables.Select<XlsxTable, string>(t => t.Name);
            if (!XlsxHelper.ValidateName("table", value, oldname, tableNames, out string message))
                throw new ArgumentException(message, nameof(value));

            _name = value;

            // Some totals row formula depend on the table name. Update them.
            if ((_fieldNames?.Count ?? 0) > 0)
                this.Fields.ForEach(f => (f as XlsxTableField).UpdateTableFieldTotalsRowFormula());

            if (!string.IsNullOrWhiteSpace(oldname) && !string.Equals(oldname, _name, StringComparison.OrdinalIgnoreCase))
            {
                Worksheet.Tables.Add(this);
                if (Worksheet.Tables.Contains(oldname))
                    Worksheet.Tables.Remove(oldname);
            }
        }
    }

    public bool ShowTotalsRow
    {
        get { return _showTotalsRow; }
        set
        {
            if (value && !_showTotalsRow)
                InsertRowsBelow(1);
            else if (!value && _showTotalsRow)
                TotalsRow().Delete();

            _showTotalsRow = value;

            // Invalidate fields' columns
            this.Fields.Cast<XlsxTableField>().ForEach(f => f.Column = null);

            if (_showTotalsRow)
            {
                AutoFilter.Range = Worksheet.Range(
                    RangeAddress.FirstAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber - 1, RangeAddress.LastAddress.ColumnNumber);
            }
            else
                AutoFilter.Range = Worksheet.Range(RangeAddress);
        }
    }

    public IXlsxRangeRow HeadersRow()
    {
        return HeadersRow(true);
    }

    internal XlsxRangeRow HeadersRow(bool scanForNewFieldsNames)
    {
        if (!ShowHeaderRow) return null;

        if (scanForNewFieldsNames)
        {
            _ = FieldNames;
        }

        return FirstRow();
    }

    public IXlsxRangeRow TotalsRow()
    {
        return ShowTotalsRow ? LastRow() : null;
    }

    public IXlsxTableField Field(string fieldName)
    {
        return Field(GetFieldIndex(fieldName));
    }

    IXlsxTableField IXlsxTable.Field(int fieldIndex)
    {
        return Field(fieldIndex);
    }

    internal XlsxTableField Field(int fieldIndex)
    {
        return FieldNames.Values.First(f => f.Index == fieldIndex);
    }

    IEnumerable<IXlsxTableField> IXlsxTable.Fields => Fields;

    internal IEnumerable<XlsxTableField> Fields
    {
        get
        {
            int columnCount = ColumnCount();
            for (int co = 0; co < columnCount; co++)
                yield return Field(co);
        }
    }

    public IXlsxTable Resize(IXlsxRangeAddress rangeAddress)
    {
        return Resize(Worksheet.Range(RangeAddress));
    }

    public IXlsxTable Resize(string rangeAddress)
    {
        return Resize(Worksheet.Range(rangeAddress));
    }

    public IXlsxTable Resize(IXlsxCell firstCell, IXlsxCell lastCell)
    {
        return Resize(Worksheet.Range(firstCell, lastCell));
    }

    public IXlsxTable Resize(string firstCellAddress, string lastCellAddress)
    {
        return Resize(Worksheet.Range(firstCellAddress, lastCellAddress));
    }

    public IXlsxTable Resize(IXlsxAddress firstCellAddress, IXlsxAddress lastCellAddress)
    {
        return Resize(Worksheet.Range(firstCellAddress, lastCellAddress));
    }

    public IXlsxTable Resize(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
    {
        return Resize(Worksheet.Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn));
    }

    public IXlsxTable Resize(IXlsxRange range)
    {
        if (!this.ShowHeaderRow)
            throw new NotImplementedException("Resizing of tables with no headers not supported yet.");

        if (this.Worksheet != range.Worksheet)
            throw new InvalidOperationException("You cannot resize a table to a range on a different sheet.");

        var totalsRowChanged = this.ShowTotalsRow ? range.LastRow().RowNumber() - this.TotalsRow().RowNumber() : 0;
        var oldTotalsRowNumber = this.ShowTotalsRow ? this.TotalsRow().RowNumber() : -1;

        Dictionary<string, XlsxTableField>.KeyCollection existingHeaders = this.FieldNames.Keys;
        var newHeaders = new HashSet<string>();

        // Force evaluation of f.Column field
        IXlsxRangeColumn[] tempArray = [.. this.Fields.Select(f => f.Column)];

        IXlsxRangeRow firstRow = range.Row(1);
        if (!firstRow.FirstCell().Address.Equals(this.HeadersRow().FirstCell().Address)
            || !firstRow.LastCell().Address.Equals(this.HeadersRow().LastCell().Address))
        {
            _uniqueNames.Clear();
            var co = 1;
            foreach (IXlsxCell c in firstRow.Cells())
            {
                if (c.IsEmpty(XlsxCellsUsedOptions.Contents))
                    c.Value = GetUniqueName("Column", co, true);

                var header = c.GetString();
                _uniqueNames.Add(header);

                if (!existingHeaders.Contains(header))
                    newHeaders.Add(header);

                co++;
            }
        }

        if (totalsRowChanged < 0)
        {
            range.Rows(r => r.RowNumber().Equals(this.TotalsRow().RowNumber() + totalsRowChanged)).Single().InsertRowsAbove(1);
            range = Worksheet.Range(range.FirstCell(), range.LastCell().CellAbove());
            oldTotalsRowNumber++;
        }
        else if (totalsRowChanged > 0)
        {
            this.TotalsRow().RowBelow(totalsRowChanged + 1).InsertRowsAbove(1);
            this.TotalsRow().AsRange().Delete(XlsxShiftDeletedCells.ShiftCellsUp);
        }

        this.RangeAddress = (XlsxRangeAddress)range.RangeAddress;
        RescanFieldNames();

        if (this.ShowTotalsRow)
        {
            foreach (XlsxTableField f in this._fieldNames.Values)
            {
                var fieldColumn = f.Index + 1;
                IXlsxCell c = this.TotalsRow().Cell(fieldColumn);
                if (!c.IsEmpty() && newHeaders.Contains(f.Name))
                {
                    f.TotalsRowLabel = c.GetFormattedString();
                }
            }

            if (totalsRowChanged != 0)
            {
                foreach (XlsxTableField f in this._fieldNames.Values.Cast<XlsxTableField>())
                {
                    f.UpdateTableFieldTotalsRowFormula();
                    var fieldColumn = f.Index + 1;
                    IXlsxCell c = this.TotalsRow().Cell(fieldColumn);
                    if (!string.IsNullOrWhiteSpace(f.TotalsRowLabel))
                    {
                        //Remove previous row's label
                        XlsxCell oldTotalsCell = Worksheet.Cell(oldTotalsRowNumber, f.Column.ColumnNumber());
                        if (oldTotalsCell.Value.Equals(f.TotalsRowLabel))
                            oldTotalsCell.Value = Blank.Value;
                    }

                    if (!string.IsNullOrEmpty(f.TotalsRowLabel))
                        c.SetValue(f.TotalsRowLabel);
                }
            }
        }

        return this;
    }

    public IXlsxTable SetEmphasizeFirstColumn()
    {
        EmphasizeFirstColumn = true;
        return this;
    }

    public IXlsxTable SetEmphasizeFirstColumn(bool value)
    {
        EmphasizeFirstColumn = value;
        return this;
    }

    public IXlsxTable SetEmphasizeLastColumn()
    {
        EmphasizeLastColumn = true;
        return this;
    }

    public IXlsxTable SetEmphasizeLastColumn(bool value)
    {
        EmphasizeLastColumn = value;
        return this;
    }

    public IXlsxTable SetShowRowStripes()
    {
        ShowRowStripes = true;
        return this;
    }

    public IXlsxTable SetShowRowStripes(bool value)
    {
        ShowRowStripes = value;
        return this;
    }

    public IXlsxTable SetShowColumnStripes()
    {
        ShowColumnStripes = true;
        return this;
    }

    public IXlsxTable SetShowColumnStripes(bool value)
    {
        ShowColumnStripes = value;
        return this;
    }

    public IXlsxTable SetShowTotalsRow()
    {
        ShowTotalsRow = true;
        return this;
    }

    public IXlsxTable SetShowTotalsRow(bool value)
    {
        ShowTotalsRow = value;
        return this;
    }

    public IXlsxTable SetShowAutoFilter()
    {
        ShowAutoFilter = true;
        return this;
    }

    public IXlsxTable SetShowAutoFilter(bool value)
    {
        ShowAutoFilter = value;
        return this;
    }

    public new IXlsxRange Sort(string columnsToSortBy, XlsxSortOrder sortOrder = XlsxSortOrder.Ascending,
                             bool matchCase = false, bool ignoreBlanks = true)
    {
        var toSortBy = new StringBuilder();
        foreach (string coPairTrimmed in columnsToSortBy.Split(',').Select(coPair => coPair.Trim()))
        {
            string coString;
            string order;
            if (coPairTrimmed.Contains(' '))
            {
                var pair = coPairTrimmed.Split(' ');
                coString = pair[0];
                order = pair[1];
            }
            else
            {
                coString = coPairTrimmed;
                order = sortOrder == XlsxSortOrder.Ascending ? "ASC" : "DESC";
            }

            if (!int.TryParse(coString, out int co))
                co = Field(coString).Index + 1;

            if (toSortBy.Length > 0)
                toSortBy.Append(',');

            toSortBy.Append(co);
            toSortBy.Append(' ');
            toSortBy.Append(order);
        }
        return DataRange.Sort(toSortBy.ToString(), sortOrder, matchCase, ignoreBlanks);
    }

    public new IXlsxTable Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    IXlsxAutoFilter IXlsxTable.AutoFilter
    {
        get { return AutoFilter; }
    }

    #endregion IXlsxTable Members

    private void InitializeValues(bool setAutofilter)
    {
        ShowRowStripes = true;
        _showHeaderRow = true;
        Theme = XlsxTableTheme.TableStyleMedium2;
        if (setAutofilter)
            InitializeAutoFilter();

        if (RowCount() == 1)
            InsertRowsBelow(1);
    }

    public void InitializeAutoFilter()
    {
        ShowAutoFilter = true;
    }

    internal void OnAddedToTables()
    {
        _uniqueNames = [];
        int co = 1;
        foreach (IXlsxCell c in Row(1).Cells())
        {
            // Be careful here. Fields names may actually be whitespace, but not empty
            if (c.IsEmpty(XlsxCellsUsedOptions.Contents))
                (c as XlsxCell).SetValue(GetUniqueName("Column", co, true), false, false);
            _uniqueNames.Add(c.GetString());
            co++;
        }
    }

    private static Dictionary<string, XlsxTableField> CreateFieldNames()
    {
        return new Dictionary<string, XlsxTableField>(StringComparer.CurrentCultureIgnoreCase);
    }

    private string GetUniqueName(string originalName, int initialOffset, bool enforceOffset)
    {
        string name = string.Concat(originalName, enforceOffset ? initialOffset.ToInvariantString() : string.Empty);
        if (_uniqueNames?.Contains(name) ?? false)
        {
            int i = initialOffset;
            name = originalName + i.ToInvariantString();
            while (_uniqueNames.Contains(name))
            {
                i++;
                name = originalName + i.ToInvariantString();
            }
        }

        return name;
    }

    public int GetFieldIndex(string name)
    {
        // There is a discrepancy in the way headers with line breaks are stored.
        // The entry in the table definition will contain \r\n
        // but the shared string value of the actual cell will contain only \n
        name = name.Replace("\r\n", "\n");
        if (FieldNames.TryGetValue(name, out XlsxTableField tableField))
            return tableField.Index;

        throw new ArgumentOutOfRangeException("The header row doesn't contain field name '" + name + "'.");
    }

    internal bool _showHeaderRow;

    public bool ShowHeaderRow
    {
        get { return _showHeaderRow; }
        set
        {
            if (_showHeaderRow == value) return;

            if (_showHeaderRow)
            {
                IXlsxRangeRow headersRow = HeadersRow();
                _uniqueNames = [];
                int co = 1;
                foreach (IXlsxCell c in headersRow.Cells())
                {
                    if (string.IsNullOrWhiteSpace(c.GetString()))
                        c.Value = GetUniqueName("Column", co, true);
                    _uniqueNames.Add(c.GetString());
                    co++;
                }

                headersRow.Clear();
                RangeAddress = new XlsxRangeAddress(
                    new XlsxAddress(Worksheet, RangeAddress.FirstAddress.RowNumber + 1,
                                  RangeAddress.FirstAddress.ColumnNumber,
                                  RangeAddress.FirstAddress.FixedRow,
                                  RangeAddress.FirstAddress.FixedColumn),
                    RangeAddress.LastAddress);
            }
            else
            {
                XlsxRange asRange = Worksheet.Range(
                    RangeAddress.FirstAddress.RowNumber - 1,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber);
                XlsxRangeRow firstRow = asRange.FirstRow();
                IXlsxRangeRow rangeRow;
                if (firstRow.IsEmpty(XlsxCellsUsedOptions.All))
                {
                    rangeRow = firstRow;
                    RangeAddress = new XlsxRangeAddress(
                        new XlsxAddress(Worksheet,
                            RangeAddress.FirstAddress.RowNumber - 1,
                            RangeAddress.FirstAddress.ColumnNumber,
                            RangeAddress.FirstAddress.FixedRow,
                            RangeAddress.FirstAddress.FixedColumn),
                        RangeAddress.LastAddress);
                }
                else
                {
                    XlsxAddress fAddress = RangeAddress.FirstAddress;
                    //var lAddress = RangeAddress.LastAddress;

                    rangeRow = firstRow.InsertRowsBelow(1, false).First();

                    RangeAddress = new XlsxRangeAddress(
                        fAddress,
                        RangeAddress.LastAddress);
                }

                int co = 1;
                foreach (var name in FieldNames.Values.Select(f => f.Name))
                {
                    rangeRow.Cell(co).SetValue(name);
                    co++;
                }
            }

            _showHeaderRow = value;

            // Invalidate fields' columns
            this.Fields.Cast<XlsxTableField>().ForEach(f => f.Column = null);
        }
    }

    public IXlsxTable SetShowHeaderRow()
    {
        return SetShowHeaderRow(true);
    }

    public IXlsxTable SetShowHeaderRow(bool value)
    {
        ShowHeaderRow = value;
        return this;
    }

    public void ExpandTableRows(int rows)
    {
        RangeAddress = new XlsxRangeAddress(
            RangeAddress.FirstAddress,
            new XlsxAddress(Worksheet, RangeAddress.LastAddress.RowNumber + rows,
                                     RangeAddress.LastAddress.ColumnNumber,
                                     RangeAddress.LastAddress.FixedRow,
                                     RangeAddress.LastAddress.FixedColumn));
    }

    public override XlsxRangeColumn Column(int columnNumber)
    {
        XlsxRangeColumn column = base.Column(columnNumber);
        column.Table = this;
        return column;
    }

    public override XlsxRangeColumn Column(string columnName)
    {
        XlsxRangeColumn column = base.Column(columnName);
        column.Table = this;
        return column;
    }

    public override IXlsxRangeColumns Columns(int firstColumn, int lastColumn)
    {
        IXlsxRangeColumns columns = base.Columns(firstColumn, lastColumn);
        columns.Cast<XlsxRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    public override IXlsxRangeColumns Columns(Func<IXlsxRangeColumn, bool> predicate = null)
    {
        IXlsxRangeColumns columns = base.Columns(predicate);
        columns.Cast<XlsxRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    public override IXlsxRangeColumns Columns(string columns)
    {
        IXlsxRangeColumns cols = base.Columns(columns);
        cols.Cast<XlsxRangeColumn>().ForEach(column => column.Table = this);
        return cols;
    }

    public override IXlsxRangeColumns Columns(string firstColumn, string lastColumn)
    {
        IXlsxRangeColumns columns = base.Columns(firstColumn, lastColumn);
        columns.Cast<XlsxRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    internal override XlsxRangeColumns ColumnsUsed(XlsxCellsUsedOptions options, Func<IXlsxRangeColumn, bool> predicate = null)
    {
        XlsxRangeColumns columns = base.ColumnsUsed(options, predicate);
        columns.Cast<XlsxRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    internal override XlsxRangeColumns ColumnsUsed(Func<IXlsxRangeColumn, bool> predicate = null)
    {
        XlsxRangeColumns columns = base.ColumnsUsed(predicate);
        columns.Cast<XlsxRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    IXlsxPivotTable IXlsxRangeBase.CreatePivotTable(IXlsxCell targetCell, string name)
    {
        return CreatePivotTable(targetCell, name);
    }

    internal new XlsxPivotTable CreatePivotTable(IXlsxCell targetCell, string name)
    {
        return (XlsxPivotTable)targetCell.Worksheet.PivotTables.Add(name, targetCell, this);
    }

    public IEnumerable<dynamic> AsDynamicEnumerable()
    {
        foreach (IXlsxTableRow row in this.DataRange.Rows())
        {
            dynamic expando = new ExpandoObject();
            foreach (XlsxTableField f in this.Fields)
            {
                XlsxCellValue value = row.Cell(f.Index + 1).Value;
                // ExpandoObject supports IDictionary so we can extend it like this
                var expandoDict = expando as IDictionary<string, object>;
                expandoDict[f.Name] = value;
            }

            yield return expando;
        }
    }

    public DataTable AsNativeDataTable()
    {
        var table = new DataTable(this.Name);

        foreach (XlsxTableField f in Fields.Cast<XlsxTableField>())
        {
            Type type = typeof(object);
            if (f.IsConsistentDataType())
            {
                IXlsxCell c = f.Column.Cells().Skip(this.ShowHeaderRow ? 1 : 0).First();
                switch (c.DataType)
                {
                    case XlsxDataType.Text:
                        type = typeof(string);
                        break;

                    case XlsxDataType.Boolean:
                        type = typeof(bool);
                        break;

                    case XlsxDataType.DateTime:
                        type = typeof(DateTime);
                        break;

                    case XlsxDataType.TimeSpan:
                        type = typeof(TimeSpan);
                        break;

                    case XlsxDataType.Number:
                        type = typeof(double);
                        break;
                }
            }

            table.Columns.Add(f.Name, type);
        }

        foreach (IXlsxTableRow row in this.DataRange.Rows())
        {
            DataRow dr = table.NewRow();

            foreach (XlsxTableField f in this.Fields)
            {
                dr[f.Name] = row.Cell(f.Index + 1).Value.ToObject();
            }

            table.Rows.Add(dr);
        }

        return table;
    }

    public IXlsxTable CopyTo(IXlsxWorksheet targetSheet)
    {
        return CopyTo((XlsxWorksheet)targetSheet);
    }

    internal IXlsxTable CopyTo(XlsxWorksheet targetSheet, bool copyData = true)
    {
        if (targetSheet == Worksheet)
            throw new InvalidOperationException("Cannot copy table to the worksheet it already belongs to.");

        XlsxRange targetRange = targetSheet.Range(RangeAddress.WithoutWorksheet());
        if (copyData)
            RangeUsed().CopyTo(targetRange);
        else
            HeadersRow().CopyTo(targetRange.FirstRow());

        string tableName = Name;
        var newTable = (XlsxTable)targetSheet.Table(targetRange, tableName, true);

        newTable.RelId = null;
        newTable.EmphasizeFirstColumn = EmphasizeFirstColumn;
        newTable.EmphasizeLastColumn = EmphasizeLastColumn;
        newTable.ShowRowStripes = ShowRowStripes;
        newTable.ShowColumnStripes = ShowColumnStripes;
        newTable.ShowAutoFilter = ShowAutoFilter;
        newTable.Theme = Theme;
        newTable._showTotalsRow = ShowTotalsRow;

        int fieldCount = ColumnCount();
        for (int f = 0; f < fieldCount; f++)
        {
            var tableField = newTable.Field(f) as XlsxTableField;
            var tField = Field(f) as XlsxTableField;
            tableField.Index = tField.Index;
            tableField.Name = tField.Name;
            tableField.totalsRowLabel = tField.totalsRowLabel;
            tableField.totalsRowFunction = tField.totalsRowFunction;
        }
        return newTable;
    }

    #region Append and replace data

    [RequiresUnreferencedCode("Table data from an IEnumerable is written with reflection over the item type and its members. Use the DataTable overloads or write XlsxCellValue values explicitly to stay trim-compatible.")]
    public IXlsxRange AppendData(IEnumerable data, bool propagateExtraColumns = false)
    {
        return AppendData(data, transpose: false, propagateExtraColumns: propagateExtraColumns);
    }

    [RequiresUnreferencedCode("Table data from an IEnumerable is written with reflection over the item type and its members. Use the DataTable overloads or write XlsxCellValue values explicitly to stay trim-compatible.")]
    public IXlsxRange AppendData(IEnumerable data, bool transpose, bool propagateExtraColumns = false)
    {
        var castedData = data?.Cast<object>().ToArray() ?? [];
        if (castedData.Length == 0 || data is string)
            return null;

        return AppendDataCore(castedData.Length, cell => cell.InsertData(castedData, transpose));
    }

    public IXlsxRange AppendData(DataTable dataTable, bool propagateExtraColumns = false)
    {
        if (dataTable.Rows.Count == 0)
            return null;

        return AppendDataCore(dataTable.Rows.Count, cell => cell.InsertData(dataTable));
    }

    [RequiresUnreferencedCode("Table data from an IEnumerable is written with reflection over the item type and its members. Use the DataTable overloads or write XlsxCellValue values explicitly to stay trim-compatible.")]
    public IXlsxRange AppendData<T>(IEnumerable<T> data, bool propagateExtraColumns = false)
    {
        T[] materializedData = data?.ToArray() ?? [];
        if (materializedData.Length == 0 || data is string)
            return null;

        return AppendDataCore(materializedData.Length, cell => cell.InsertData(materializedData));
    }

    [RequiresUnreferencedCode("Table data from an IEnumerable is written with reflection over the item type and its members. Use the DataTable overloads or write XlsxCellValue values explicitly to stay trim-compatible.")]
    public IXlsxRange ReplaceData(IEnumerable data, bool propagateExtraColumns = false)
    {
        return ReplaceData(data, transpose: false, propagateExtraColumns: propagateExtraColumns);
    }

    [RequiresUnreferencedCode("Table data from an IEnumerable is written with reflection over the item type and its members. Use the DataTable overloads or write XlsxCellValue values explicitly to stay trim-compatible.")]
    public IXlsxRange ReplaceData(IEnumerable data, bool transpose, bool propagateExtraColumns = false)
    {
        var castedData = data?.Cast<object>().ToArray() ?? [];
        if (castedData.Length == 0 || data is string)
            throw new InvalidOperationException("Cannot replace table data with empty enumerable.");

        return ReplaceDataCore(castedData.Length, propagateExtraColumns, cell => cell.InsertData(castedData, transpose));
    }

    public IXlsxRange ReplaceData(DataTable dataTable, bool propagateExtraColumns = false)
    {
        if (dataTable.Rows.Count == 0)
            throw new InvalidOperationException("Cannot replace table data with empty enumerable.");

        return ReplaceDataCore(dataTable.Rows.Count, propagateExtraColumns, cell => cell.InsertData(dataTable));
    }

    [RequiresUnreferencedCode("Table data from an IEnumerable is written with reflection over the item type and its members. Use the DataTable overloads or write XlsxCellValue values explicitly to stay trim-compatible.")]
    public IXlsxRange ReplaceData<T>(IEnumerable<T> data, bool propagateExtraColumns = false)
    {
        T[] materializedData = data?.ToArray() ?? [];
        if (materializedData.Length == 0 || data is string)
            throw new InvalidOperationException("Cannot replace table data with empty enumerable.");

        return ReplaceDataCore(materializedData.Length, propagateExtraColumns, cell => cell.InsertData(materializedData));
    }

    private IXlsxRange AppendDataCore(int numberOfNewRows, Func<IXlsxCell, IXlsxRange> insertData)
    {
        IXlsxTableRow lastRowOfOldRange = this.DataRange.LastRow();
        lastRowOfOldRange.InsertRowsBelow(numberOfNewRows);
        this.Fields.Cast<XlsxTableField>().ForEach(f => f.Column = null);

        IXlsxRange insertedRange = insertData(lastRowOfOldRange.RowBelow().FirstCell());

        PropagateExtraColumns(insertedRange.ColumnCount(), lastRowOfOldRange.RowNumber());

        return insertedRange;
    }

    private IXlsxRange ReplaceDataCore(int numberOfRows, bool propagateExtraColumns, Func<IXlsxCell, IXlsxRange> insertData)
    {
        var firstDataRowNumber = this.DataRange.FirstRow().RowNumber();
        var lastDataRowNumber = this.DataRange.LastRow().RowNumber();

        // Resize table
        var sizeDifference = numberOfRows - this.DataRange.RowCount();
        if (sizeDifference > 0)
            this.DataRange.LastRow().InsertRowsBelow(sizeDifference);
        else if (sizeDifference < 0)
        {
            this.DataRange.Rows
            (
                lastDataRowNumber + sizeDifference + 1 - firstDataRowNumber + 1,
                lastDataRowNumber - firstDataRowNumber + 1
            )
            .Delete();

            // No propagation needed when reducing the number of rows
            propagateExtraColumns = false;
        }

        if (sizeDifference != 0)
            // Invalidate table fields' columns
            this.Fields.Cast<XlsxTableField>().ForEach(f => f.Column = null);

        IXlsxRange replacedRange = insertData(this.DataRange.FirstCell());

        if (propagateExtraColumns)
            PropagateExtraColumns(replacedRange.ColumnCount(), lastDataRowNumber);

        return replacedRange;
    }
    private void PropagateExtraColumns(int numberOfNonExtraColumns, int previousLastDataRow)
    {
        for (var i = numberOfNonExtraColumns; i < this.Fields.Count(); i++)
        {
            XlsxTableField field = this.Field(i);

            XlsxCell cell = this.Worksheet.Cell(previousLastDataRow, field.Column.ColumnNumber());
            field.Column.Cells(c => c.Address.RowNumber > previousLastDataRow)
                .ForEach(c =>
                {
                    if (cell.HasFormula)
                        c.FormulaR1C1 = cell.FormulaR1C1;
                    else
                        c.Value = cell.Value;
                });
        }
    }

    /// <summary>
    /// Update headers fields and totals fields by data from the cells. Do not add a new fields or names.
    /// </summary>
    /// <param name="refreshArea">Area that contains cells with changed values that might affect header and totals fields.</param>
    internal void RefreshFieldsFromCells(Area refreshArea)
    {
        Area tableArea = Area;
        if (ShowTotalsRow)
        {
            Area totalsRow = tableArea.SliceFromBottom(1);
            Area? intersection = totalsRow.Intersect(refreshArea);
            if (intersection is not null)
            {
                var totalsRowNumber = totalsRow.BottomRow;
                ValueSlice valueSlice = Worksheet.Internals.CellsCollection.ValueSlice;
                for (var column = intersection.Value.LeftColumn; column <= intersection.Value.RightColumn; ++column)
                {
                    var fieldIndex = column - totalsRow.LeftColumn;
                    XlsxTableField field = Field(fieldIndex);
                    XlsxCellValue value = valueSlice.GetCellValue(new Point(totalsRowNumber, column));

                    // Convert value to text, because Excel always converts values to text when replacing totals row.
                    field.TotalsRowLabel = value.ToString(CultureInfo.CurrentCulture);
                }
            }
        }

        if (ShowHeaderRow)
        {
            Area headersRow = Area.SliceFromTop(1);
            Area? intersection = headersRow.Intersect(refreshArea);
            if (intersection is not null)
            {
                var headersRowNumber = headersRow.TopRow;
                ValueSlice valueSlice = Worksheet.Internals.CellsCollection.ValueSlice;
                for (var column = intersection.Value.LeftColumn; column <= intersection.Value.RightColumn; ++column)
                {
                    var fieldIndex = column - headersRow.LeftColumn;
                    XlsxTableField field = Field(fieldIndex);
                    XlsxCellValue value = valueSlice.GetCellValue(new Point(headersRowNumber, column));

                    // Convert to text, because headers row of a table can be only
                    // string in OOXML and Excel converts it to string as well.
                    field.Name = value.ToString(CultureInfo.CurrentCulture);
                }
            }
        }
    }

    #endregion Append and replace data
}
