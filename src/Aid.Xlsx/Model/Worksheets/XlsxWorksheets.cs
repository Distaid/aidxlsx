using System.Collections;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Diagnostics;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.DefinedNames;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Model.Worksheets;

internal class XlsxWorksheets(XlsxWorkbook workbook) : IXlsxWorksheets, IEnumerable<XlsxWorksheet>
{
    private readonly XlsxWorkbook _workbook = workbook;
    private readonly Dictionary<string, XlsxWorksheet> _worksheets = [with(XlsxHelper.SheetComparer)];
    internal ICollection<String> Deleted { get; private set; } = new HashSet<String>();

    /// <summary>
    /// SheetId that will be assigned to next created sheet.
    /// </summary>
    private UInt32 _nextSheetId = 1;

    #region Constructor

    #endregion Constructor

    #region IEnumerable<XlsxWorksheet> Members

    public IEnumerator<XlsxWorksheet> GetEnumerator()
    {
        return ((IEnumerable<XlsxWorksheet>)_worksheets.Values).GetEnumerator();
    }

    #endregion IEnumerable<XlsxWorksheet> Members

    #region IXlsxWorksheets Members

    public int Count
    {
        [DebuggerStepThrough]
        get { return _worksheets.Count; }
    }

    public Boolean Contains(String sheetName)
    {
        return _worksheets.ContainsKey(sheetName);
    }

    bool IXlsxWorksheets.TryGetWorksheet(string sheetName, [NotNullWhen(true)] out IXlsxWorksheet? worksheet)
    {
        if (TryGetWorksheet(sheetName, out var foundSheet))
        {
            worksheet = foundSheet;
            return true;
        }

        worksheet = null;
        return false;
    }

    internal bool TryGetWorksheet(string sheetName, [NotNullWhen(true)] out XlsxWorksheet? worksheet)
    {
        return _worksheets.TryGetValue(sheetName, out worksheet);
    }

    public IXlsxWorksheet Worksheet(string sheetName)
    {
        if (!_worksheets.TryGetValue(sheetName, out var foundSheet))
            throw new KeyNotFoundException($"There isn't a worksheet named '{sheetName}'.");

        return foundSheet;
    }

    public IXlsxWorksheet Worksheet(Int32 position)
    {
        int wsCount = _worksheets.Values.Count(w => w.Position == position);
        if (wsCount == 0)
            throw new ArgumentException("There isn't a worksheet associated with that position.");

        if (wsCount > 1)
        {
            throw new ArgumentException(
                "Can't retrieve a worksheet because there are multiple worksheets associated with that position.");
        }

        return _worksheets.Values.Single(w => w.Position == position);
    }

    public IXlsxWorksheet Add()
    {
        return Add(GetNextWorksheetName());
    }

    public IXlsxWorksheet Add(Int32 position)
    {
        return Add(GetNextWorksheetName(), position);
    }

    public IXlsxWorksheet Add(String sheetName)
    {
        var sheet = new XlsxWorksheet(sheetName, _workbook, GetNextSheetId());
        Add(sheetName, sheet);
        sheet._position = _worksheets.Count + _workbook.UnsupportedSheets.Count;
        return sheet;
    }

    public IXlsxWorksheet Add(String sheetName, Int32 position)
    {
        return Add(sheetName, position, GetNextSheetId());
    }

    internal XlsxWorksheet Add(String sheetName, Int32 position, UInt32 sheetId)
    {
        _worksheets.Values.Where(w => w._position >= position).ForEach(w => w._position += 1);
        _workbook.UnsupportedSheets.Where(w => w.Position >= position).ForEach(w => w.Position += 1);

        // If the loaded sheetId is greater than current, just make sure our next sheetId is even bigger.
        _nextSheetId = Math.Max(_nextSheetId, sheetId + 1);
        var sheet = new XlsxWorksheet(sheetName, _workbook, sheetId);
        Add(sheetName, sheet);
        sheet._position = position;
        return sheet;
    }

    private void Add(String sheetName, XlsxWorksheet sheet)
    {
        if (_worksheets.ContainsKey(sheetName))
            throw new ArgumentException(String.Format("A worksheet with the same name ({0}) has already been added.", sheetName), nameof(sheetName));

        _worksheets.Add(sheetName, sheet);

        _workbook.NotifyWorksheetAdded(sheet);
    }

    public void Delete(String sheetName)
    {
        Delete(_worksheets[sheetName].Position);
    }

    public void Delete(Int32 position)
    {
        int wsCount = _worksheets.Values.Count(w => w.Position == position);
        if (wsCount == 0)
            throw new ArgumentException("There isn't a worksheet associated with that index.");

        if (wsCount > 1)
            throw new ArgumentException(
                "Can't delete the worksheet because there are multiple worksheets associated with that index.");

        var ws = _worksheets.Values.Single(w => w.Position == position);
        if (!String.IsNullOrWhiteSpace(ws.RelId) && !Deleted.Contains(ws.RelId))
            Deleted.Add(ws.RelId);

        _worksheets.RemoveAll(w => w.Position == position);
        _worksheets.Values.Where(w => w.Position > position).ForEach(w => w._position -= 1);
        _workbook.UnsupportedSheets.Where(w => w.Position > position).ForEach(w => w.Position -= 1);

        ws.Cleanup();
    }

    IEnumerator<IXlsxWorksheet> IEnumerable<IXlsxWorksheet>.GetEnumerator()
    {
        return _worksheets.Values.Cast<IXlsxWorksheet>().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public IXlsxWorksheet Add(DataTable dataTable)
    {
        return Add(dataTable, dataTable.TableName);
    }

    public IXlsxWorksheet Add(DataTable dataTable, String sheetName)
    {
        return Add(dataTable, sheetName, TableNameGenerator.GetNewTableName(_workbook));
    }

    public IXlsxWorksheet Add(DataTable dataTable, String sheetName, string tableName)
    {
        var ws = Add(sheetName);
        ws.Cell(1, 1).InsertTable(dataTable, tableName);
        return ws;
    }

    public void Add(DataSet dataSet)
    {
        foreach (DataTable t in dataSet.Tables)
            Add(t);
    }

    #endregion IXlsxWorksheets Members

    public void Rename(String oldSheetName, String newSheetName)
    {
        if (String.IsNullOrWhiteSpace(oldSheetName) || !_worksheets.TryGetValue(oldSheetName, out XlsxWorksheet? ws)) return;

        if (!oldSheetName.Equals(newSheetName, StringComparison.OrdinalIgnoreCase)
            && _worksheets.ContainsKey(newSheetName))
            throw new ArgumentException(String.Format("A worksheet with the same name ({0}) has already been added.", newSheetName), nameof(newSheetName));

        _worksheets.Remove(oldSheetName);
        Add(newSheetName, ws);

        foreach (var listener in GetWorkbookListeners())
            listener.OnSheetRenamed(oldSheetName, newSheetName);
    }

    #region Private members

    private IEnumerable<IWorkbookListener> GetWorkbookListeners()
    {
        // All components that should be updated when sheet is added/removed or renamed should
        // be enumerated here.
        yield return _workbook.CalcEngine;

        foreach (var sheet in _worksheets.Values)
        {
            yield return sheet.Internals.CellsCollection;
        }

        foreach (var definedName in _workbook.DefinedNamesInternal)
            yield return definedName;

        foreach (var sheet in _worksheets.Values)
        {
            foreach (var definedName in sheet.DefinedNames)
            {
                yield return definedName;
            }
        }
    }

    private String GetNextWorksheetName()
    {
        var worksheetNumber = this.Count + 1;
        var sheetName = $"Sheet{worksheetNumber}";
        while (_worksheets.Values.Any(p => p.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
        {
            worksheetNumber++;
            sheetName = $"Sheet{worksheetNumber}";
        }
        return sheetName;
    }

    private UInt32 GetNextSheetId() => _nextSheetId++;

    #endregion Private members
}
