using System.Diagnostics;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Tables;

[DebuggerDisplay("{Name}")]
internal class XlsxTableField(XlsxTable table, String name) : IXlsxTableField
{
    internal XlsxTotalsRowFunction totalsRowFunction;
    internal String? totalsRowLabel;
    private readonly XlsxTable table = table;

    private IXlsxRangeColumn? _column;
    private Int32 index;
    private String name = name;

    public IXlsxRangeColumn Column
    {
        get
        {
            _column ??= table.AsRange().Column(this.Index + 1);
            return _column;
        }
        internal set
        {
            _column = value;
        }
    }

    public IXlsxCells DataCells
    {
        get
        {
            return Column.Cells(c =>
            {
                if (table.ShowHeaderRow && c.Equals(HeaderCell))
                    return false;
                if (table.ShowTotalsRow && c.Equals(TotalsCell))
                    return false;
                return true;
            });
        }
    }

    public IXlsxCell? HeaderCell
    {
        get
        {
            if (!table.ShowHeaderRow)
                return null;

            return Column.FirstCell();
        }
    }

    public Int32 Index
    {
        get { return index; }
        internal set
        {
            if (index == value) return;
            index = value;
            _column = null;
        }
    }

    public String Name
    {
        get
        {
            return name;
        }
        set
        {
            if (name == value) return;

            if (table.ShowHeaderRow)
                ((XlsxCell)table.HeadersRow(false).Cell(Index + 1)).SetValue(value, setTableHeader: false, checkMergedRanges: true);

            table.RenameField(name, value);
            name = value;
        }
    }

    public IXlsxTable Table { get { return table; } }

    public IXlsxCell? TotalsCell
    {
        get
        {
            if (!table.ShowTotalsRow)
                return null;

            return Column.LastCell();
        }
    }

    public String TotalsRowFormulaA1
    {
        get { return table.TotalsRow().Cell(Index + 1).FormulaA1; }
        set
        {
            totalsRowFunction = XlsxTotalsRowFunction.Custom;
            table.TotalsRow().Cell(Index + 1).FormulaA1 = value;
        }
    }

    public String TotalsRowFormulaR1C1
    {
        get { return table.TotalsRow().Cell(Index + 1).FormulaR1C1; }
        set
        {
            totalsRowFunction = XlsxTotalsRowFunction.Custom;
            table.TotalsRow().Cell(Index + 1).FormulaR1C1 = value;
        }
    }

    public XlsxTotalsRowFunction TotalsRowFunction
    {
        get { return totalsRowFunction; }
        set
        {
            totalsRowFunction = value;
            UpdateTableFieldTotalsRowFormula();
        }
    }

    public String? TotalsRowLabel
    {
        get { return totalsRowLabel; }
        set
        {
            totalsRowFunction = XlsxTotalsRowFunction.None;
            ((XlsxCell)table.TotalsRow().Cell(Index + 1)).SetValue(value, setTableHeader: false, checkMergedRanges: true);
            totalsRowLabel = value;
        }
    }

    /// <summary>
    /// Dxf of header row cells of the fields column.
    /// </summary>
    internal XlsxDxfValue? HeaderFormatValue { get; set; }

    /// <summary>
    /// Dxf of total data cells of the fields column.
    /// </summary>
    internal XlsxDxfValue? DataFormatValue { get; set; }

    /// <summary>
    /// Dxf of total row cells of the fields column.
    /// </summary>
    internal XlsxDxfValue? TotalFormatValue { get; set; }

    public void Delete()
    {
        Delete(true);
    }

    internal void Delete(Boolean deleteUnderlyingRangeColumn)
    {
        var fields = table.Fields.Cast<XlsxTableField>().ToArray();

        if (deleteUnderlyingRangeColumn)
        {
            table.AsRange().ColumnQuick(this.Index + 1).Delete();
        }

        fields.Where(f => f.Index > this.Index).ForEach(f => f.Index--);
        table.FieldNames.Remove(this.Name);
    }

    public bool IsConsistentDataType()
    {
        var dataTypes = this.Column
            .Cells()
            .Skip(this.table.ShowHeaderRow ? 1 : 0)
            .Select(c => c.DataType);

        if (this.table.ShowTotalsRow)
            dataTypes = dataTypes.SkipLast();

        var distinctDataTypes = dataTypes
            .GroupBy(dt => dt)
            .Select(g => new { g.Key, Count = g.Count() });

        return distinctDataTypes.Count() == 1;
    }

    public Boolean IsConsistentFormula()
    {
        var formulas = this.Column
            .Cells()
            .Skip(this.table.ShowHeaderRow ? 1 : 0)
            .Select(c => c.FormulaR1C1);

        if (this.table.ShowTotalsRow)
            formulas = formulas.SkipLast();

        var distinctFormulas = formulas
            .GroupBy(f => f)
            .Select(g => new { g.Key, Count = g.Count() });

        return distinctFormulas.Count() == 1;
    }

    public bool IsConsistentStyle()
    {
        var styles = this.Column
            .Cells()
            .Skip(this.table.ShowHeaderRow ? 1 : 0)
            .OfType<XlsxCell>()
            .Select(c => c.Format);

        if (this.table.ShowTotalsRow)
            styles = styles.SkipLast();

        var distinctStyles = styles
            .Distinct();

        return distinctStyles.Count() == 1;
    }

    private static readonly IEnumerable<String> QuotedTableFieldCharacters = new[] { "'", "#" };

    internal void UpdateTableFieldTotalsRowFormula()
    {
        if (TotalsRowFunction != XlsxTotalsRowFunction.None && TotalsRowFunction != XlsxTotalsRowFunction.Custom)
        {
            var cell = table.TotalsRow().Cell(Index + 1);
            var formulaCode = String.Empty;
            switch (TotalsRowFunction)
            {
                case XlsxTotalsRowFunction.Sum: formulaCode = "109"; break;
                case XlsxTotalsRowFunction.Minimum: formulaCode = "105"; break;
                case XlsxTotalsRowFunction.Maximum: formulaCode = "104"; break;
                case XlsxTotalsRowFunction.Average: formulaCode = "101"; break;
                case XlsxTotalsRowFunction.Count: formulaCode = "103"; break;
                case XlsxTotalsRowFunction.CountNumbers: formulaCode = "102"; break;
                case XlsxTotalsRowFunction.StandardDeviation: formulaCode = "107"; break;
                case XlsxTotalsRowFunction.Variance: formulaCode = "110"; break;
            }

            var modifiedName = Name;
            QuotedTableFieldCharacters.ForEach(c => modifiedName = modifiedName.Replace(c, "'" + c));

            if (modifiedName.StartsWith(' ') || modifiedName.EndsWith(' '))
            {
                modifiedName = "[" + modifiedName + "]";
            }

            var prependTableName = modifiedName.Contains(' ');

            cell.FormulaA1 = $"SUBTOTAL({formulaCode},{(prependTableName ? table.Name : string.Empty)}[{modifiedName}])";
            var lastCell = table.LastRow().Cell(Index + 1);
            if (lastCell.DataType != XlsxDataType.Text)
            {
                cell.Style.NumberFormat = lastCell.Style.NumberFormat;
            }
        }
    }
}
