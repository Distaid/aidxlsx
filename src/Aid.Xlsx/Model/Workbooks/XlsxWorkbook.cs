#nullable disable

using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using static Aid.Xlsx.Model.Protection.XlsxProtectionAlgorithm;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.CustomProperties;
using Aid.Xlsx.Model.DefinedNames;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Columns;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Rows;
using Aid.Xlsx.Model.PageSetup;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.Protection;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.Rendering;
using Aid.Xlsx.SpreadsheetML.Workbooks;

namespace Aid.Xlsx.Model.Workbooks;

public enum XlsxCalculateMode
{
    Auto,
    AutoNoTable,
    Manual,
    Default
};

public enum XlsxReferenceStyle
{
    R1C1,
    A1,
    Default
};

public enum XlsxCellSetValueBehavior
{
    /// <summary>
    ///   Analyze input string and convert value. For avoid analyzing use escape symbol '
    /// </summary>
    Smart = 0,

    /// <summary>
    ///   Direct set value. If value has unsupported type - value will be stored as string returned by <see
    ///    cref = "object.ToString()" />
    /// </summary>
    Simple = 1,
}

public sealed partial class XlsxWorkbook : IXlsxWorkbook
{
    private bool _disposed;

    #region Static

    public static double DefaultRowHeight { get; private set; }
    public static double DefaultColumnWidth { get; private set; }

    public static IXlsxPageSetup DefaultPageOptions
    {
        get
        {
            var defaultPageOptions = new XlsxPageSetup(null, null)
            {
                PageOrientation = XlsxPageOrientation.Default,
                Scale = 100,
                PaperSize = XlsxPaperSize.LetterPaper,
                Margins = new XlsxMargins
                {
                    Top = 0.75,
                    Bottom = 0.5,
                    Left = 0.75,
                    Right = 0.75,
                    Header = 0.5,
                    Footer = 0.75
                },
                ScaleHFWithDocument = true,
                AlignHFWithMargins = true,
                PrintErrorValue = XlsxPrintErrorValues.Displayed,
                ShowComments = XlsxShowCommentsValues.None
            };
            return defaultPageOptions;
        }
    }

    public static IXlsxOutline DefaultOutline
    {
        get
        {
            return new XlsxOutline(null)
            {
                SummaryHLocation = XlsxOutlineSummaryHLocation.Right,
                SummaryVLocation = XlsxOutlineSummaryVLocation.Bottom
            };
        }
    }

    /// <summary>
    ///   Behavior for <see cref = "IXlsxCell.set_Value" />
    /// </summary>
    public static XlsxCellSetValueBehavior CellSetValueBehavior { get; set; }

    public static XlsxWorkbook OpenFromTemplate(string path)
    {
        return new XlsxWorkbook(path, asTemplate: true);
    }

    #endregion Static

    internal readonly List<UnsupportedSheet> UnsupportedSheets = [];

    internal IXlsxGraphicEngine GraphicEngine { get; }

    internal double DpiX { get; }

    internal double DpiY { get; }

    /// <inheritdoc cref="LoadOptions.StrictAttributeParsing"/>
    internal bool StrictAttributeParsing { get; }

    internal CancellationToken CancellationToken { get; }

    internal long MaxCalculationMemoryBytes { get; }

    internal XlsxPivotCaches PivotCachesInternal
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        private init;
    }

    internal SharedStringTable SharedStringTable
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
    } = new();

    internal XlsxWorkbookStyles Styles
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
    }

    #region Nested Type : XlsxLoadSource

    private enum XlsxLoadSource
    {
        New,
        File,
        Stream
    };

    #endregion Nested Type : XlsxLoadSource

    internal XlsxWorksheets WorksheetsInternal
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        private init;
    }

    /// <summary>
    ///   Gets an object to manipulate the worksheets.
    /// </summary>
    public IXlsxWorksheets Worksheets => WorksheetsInternal;

    internal XlsxDefinedNames DefinedNamesInternal
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
    }

    [Obsolete($"Use {nameof(DefinedNames)} instead.")]
    public IXlsxDefinedNames NamedRanges => DefinedNamesInternal;

    /// <summary>
    ///   Gets an object to manipulate this workbook's named ranges.
    /// </summary>
    public IXlsxDefinedNames DefinedNames => DefinedNamesInternal;

    /// <summary>
    ///   Gets an object to manipulate this workbook's theme.
    /// </summary>
    public IXlsxTheme Theme
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        private set
        {
            ThrowIfDisposed();
            field = value;
        }
    }

    /// <summary>
    /// All pivot caches in the workbook, whether they have a pivot table or not.
    /// </summary>
    public IXlsxPivotCaches PivotCaches => PivotCachesInternal;

    /// <inheritdoc/>
    public IXlsxStyle Style
    {
        get
        {
            ThrowIfDisposed();
            return Format;
        }
        set
        {
            ThrowIfDisposed();
            Format.SetStyle(value);
        }
    }

    /// <summary>
    ///   Gets or sets the default row height for the workbook.
    ///   <para>All new worksheets will use this row height.</para>
    /// </summary>
    public double RowHeight { get; set; }

    /// <summary>
    ///   Gets or sets the default column width for the workbook.
    ///   <para>All new worksheets will use this column width.</para>
    /// </summary>
    public double ColumnWidth { get; set; }

    /// <summary>
    ///   Gets or sets the default page options for the workbook.
    ///   <para>All new worksheets will use these page options.</para>
    /// </summary>
    public IXlsxPageSetup PageOptions
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        set
        {
            ThrowIfDisposed();
            field = value;
        }
    }

    /// <summary>
    ///   Gets or sets the default outline options for the workbook.
    ///   <para>All new worksheets will use these outline options.</para>
    /// </summary>
    public IXlsxOutline Outline
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        set
        {
            ThrowIfDisposed();
            field = value;
        }
    }

    /// <summary>
    ///   Gets or sets the workbook's properties.
    /// </summary>
    public XlsxWorkbookProperties Properties
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        set
        {
            ThrowIfDisposed();
            field = value;
        }
    }

    /// <summary>
    ///   Gets or sets the workbook's calculation mode.
    /// </summary>
    public XlsxCalculateMode CalculateMode { get; set; }

    public bool CalculationOnSave { get; set; }
    public bool ForceFullCalculation { get; set; }
    public bool FullCalculationOnLoad { get; set; }
    public bool FullPrecision { get; set; }

    /// <summary>
    ///   Gets or sets the workbook's reference style.
    /// </summary>
    public XlsxReferenceStyle ReferenceStyle { get; set; }

    public IXlsxCustomProperties CustomProperties
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        private init;
    }

    public bool ShowFormulas { get; set; }
    public bool ShowGridLines { get; set; }
    public bool ShowOutlineSymbols { get; set; }
    public bool ShowRowColHeaders { get; set; }
    public bool ShowRuler { get; set; }
    public bool ShowWhiteSpace { get; set; }
    public bool ShowZeros { get; set; }
    public bool RightToLeft { get; set; }

    public bool DefaultShowFormulas
    {
        get { return false; }
    }

    public bool DefaultShowGridLines
    {
        get { return true; }
    }

    public bool DefaultShowOutlineSymbols
    {
        get { return true; }
    }

    public bool DefaultShowRowColHeaders
    {
        get { return true; }
    }

    public bool DefaultShowRuler
    {
        get { return true; }
    }

    public bool DefaultShowWhiteSpace
    {
        get { return true; }
    }

    public bool DefaultShowZeros
    {
        get { return true; }
    }

    public IXlsxFileSharing FileSharing
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
    } = new XlsxFileSharing();

    public bool DefaultRightToLeft
    {
        get { return false; }
    }

    private void InitializeTheme()
    {
        Theme = new XlsxTheme
        {
            Text1 = XlsxColor.FromHtml("#FF000000"),
            Background1 = XlsxColor.FromHtml("#FFFFFFFF"),
            Text2 = XlsxColor.FromHtml("#FF1F497D"),
            Background2 = XlsxColor.FromHtml("#FFEEECE1"),
            Accent1 = XlsxColor.FromHtml("#FF4F81BD"),
            Accent2 = XlsxColor.FromHtml("#FFC0504D"),
            Accent3 = XlsxColor.FromHtml("#FF9BBB59"),
            Accent4 = XlsxColor.FromHtml("#FF8064A2"),
            Accent5 = XlsxColor.FromHtml("#FF4BACC6"),
            Accent6 = XlsxColor.FromHtml("#FFF79646"),
            Hyperlink = XlsxColor.FromHtml("#FF0000FF"),
            FollowedHyperlink = XlsxColor.FromHtml("#FF800080")
        };
    }

#nullable enable
    [Obsolete($"Use {nameof(DefinedName)} instead.")]
    public IXlsxDefinedName? NamedRange(string name) => DefinedName(name);

    /// <inheritdoc/>
    public IXlsxDefinedName? DefinedName(string name)
    {
        ThrowIfDisposed();
        if (name.Contains('!'))
        {
            var split = name.Split('!');
            var first = split[0];
            var wsName = first.StartsWith('\'') ? first[1..^1] : first;
            var sheetlessName = split[1];
            if (TryGetWorksheet(wsName, out XlsxWorksheet ws))
            {
                if (ws.DefinedNames.TryGetScopedValue(sheetlessName, out XlsxDefinedName? sheetDefinedName))
                    return sheetDefinedName;
            }

            name = sheetlessName;
        }

        return DefinedNamesInternal.TryGetScopedValue(name, out XlsxDefinedName? definedName) ? definedName : null;
    }
#nullable disable

    public bool TryGetWorksheet(string name, out IXlsxWorksheet worksheet)
    {
        ThrowIfDisposed();
        if (TryGetWorksheet(name, out XlsxWorksheet foundSheet))
        {
            worksheet = foundSheet;
            return true;
        }

        worksheet = default;
        return false;
    }

    internal bool TryGetWorksheet(string name, [NotNullWhen(true)] out XlsxWorksheet worksheet)
    {
        ThrowIfDisposed();
        return WorksheetsInternal.TryGetWorksheet(name, out worksheet);
    }

    public IXlsxRange RangeFromFullAddress(string rangeAddress, out IXlsxWorksheet ws)
    {
        ThrowIfDisposed();
        if (!rangeAddress.Contains('!'))
        {
            ws = null;
            return null;
        }

        var split = rangeAddress.Split('!');
        var wsName = split[0].UnescapeSheetName();
        if (TryGetWorksheet(wsName, out XlsxWorksheet sheet))
        {
            ws = sheet;
            return sheet.Range(split[1]);
        }

        ws = null;
        return null;
    }

    public IXlsxCell CellFromFullAddress(string cellAddress, out IXlsxWorksheet ws)
    {
        ThrowIfDisposed();
        if (!cellAddress.Contains('!'))
        {
            ws = null;
            return null;
        }

        var split = cellAddress.Split('!');
        var wsName = split[0].UnescapeSheetName();
        if (TryGetWorksheet(wsName, out XlsxWorksheet sheet))
        {
            ws = sheet;
            return sheet.Cell(split[1]);
        }

        ws = null;
        return null;
    }

    /// <summary>
    ///   Saves the current workbook.
    /// </summary>
    public void Save()
    {
        ThrowIfDisposed();
#if DEBUG
        Save(true, false);
#else
        Save(false, false);
#endif
    }

    /// <summary>
    ///   Saves the current workbook and optionally performs validation
    /// </summary>
    public void Save(bool validate, bool evaluateFormulae = false)
    {
        ThrowIfDisposed();
        Save(new SaveOptions
        {
            ValidatePackage = validate,
            EvaluateFormulasBeforeSaving = evaluateFormulae,
            GenerateCalculationChain = true
        });
    }

    public void Save(SaveOptions options)
    {
        ThrowIfDisposed();
        CheckForWorksheetsPresent();
        if (_loadSource == XlsxLoadSource.New)
            throw new InvalidOperationException("This is a new file. Please use one of the 'SaveAs' methods.");

        if (_loadSource == XlsxLoadSource.Stream)
        {
            CreatePackage(_originalStream, false, _spreadsheetDocumentType, options);
        }
        else
            CreatePackage(_originalFile, _spreadsheetDocumentType, options);
    }

    /// <summary>
    ///   Saves the current workbook to a file.
    /// </summary>
    public void SaveAs(string file)
    {
        ThrowIfDisposed();
#if DEBUG
        SaveAs(file, true, false);
#else
        SaveAs(file, false, false);
#endif
    }

    /// <summary>
    ///   Saves the current workbook to a file and optionally validates it.
    /// </summary>
    public void SaveAs(string file, bool validate, bool evaluateFormulae = false)
    {
        ThrowIfDisposed();
        SaveAs(file, new SaveOptions
        {
            ValidatePackage = validate,
            EvaluateFormulasBeforeSaving = evaluateFormulae,
            GenerateCalculationChain = true
        });
    }

    public void SaveAs(string file, SaveOptions options)
    {
        ThrowIfDisposed();
        CheckForWorksheetsPresent();

        var directoryName = Path.GetDirectoryName(file);
        if (!string.IsNullOrWhiteSpace(directoryName)) Directory.CreateDirectory(directoryName);

        if (_loadSource == XlsxLoadSource.New)
        {
            if (File.Exists(file))
                File.Delete(file);

            CreatePackage(file, GetWorkbookPackageType(file), options);
        }
        else if (_loadSource == XlsxLoadSource.File)
        {
            if (string.Compare(_originalFile.Trim(), file.Trim(), true) != 0)
            {
                File.Copy(_originalFile, file, true);
                File.SetAttributes(file, FileAttributes.Normal);
            }

            CreatePackage(file, GetWorkbookPackageType(file), options);
        }
        else if (_loadSource == XlsxLoadSource.Stream)
        {
            _originalStream.Position = 0;

            using FileStream fileStream = File.Create(file);
            CopyStream(_originalStream, fileStream);
            CreatePackage(fileStream, false, _spreadsheetDocumentType, options);
        }

        _loadSource = XlsxLoadSource.File;
        _originalFile = file;
        _originalStream = null;
    }

    private static WorkbookPackageType GetWorkbookPackageType(string filePath)
    {
        var extension = Path.GetExtension(filePath);

        if (string.IsNullOrEmpty(extension)) throw new ArgumentException("Empty extension is not supported.");
        extension = extension[1..].ToLowerInvariant();

        return extension switch
        {
            "xlsm" => WorkbookPackageType.MacroEnabledWorkbook,
            "xltm" => WorkbookPackageType.MacroEnabledTemplate,
            "xlsx" => WorkbookPackageType.Workbook,
            "xltx" => WorkbookPackageType.Template,
            _ => throw new ArgumentException(string.Format("Extension '{0}' is not supported. Supported extensions are '.xlsx', '.xlsm', '.xltx' and '.xltm'.", extension)),
        };
    }

    private void CheckForWorksheetsPresent()
    {
        if (Worksheets.Count == 0)
            throw new InvalidOperationException("Workbooks need at least one worksheet.");
    }

    /// <summary>
    ///   Saves the current workbook to a stream.
    /// </summary>
    public void SaveAs(Stream stream)
    {
        ThrowIfDisposed();
#if DEBUG
        SaveAs(stream, true, false);
#else
        SaveAs(stream, false, false);
#endif
    }

    /// <summary>
    ///   Saves the current workbook to a stream and optionally validates it.
    /// </summary>
    public void SaveAs(Stream stream, bool validate, bool evaluateFormulae = false)
    {
        ThrowIfDisposed();
        SaveAs(stream, new SaveOptions
        {
            ValidatePackage = validate,
            EvaluateFormulasBeforeSaving = evaluateFormulae,
            GenerateCalculationChain = true
        });
    }

    public void SaveAs(Stream stream, SaveOptions options)
    {
        ThrowIfDisposed();
        CheckForWorksheetsPresent();
        if (_loadSource == XlsxLoadSource.New)
        {
            // dm 20130422, this method or better the method SpreadsheetDocument.Create which is called
            // inside of 'CreatePackage' need a stream which CanSeek & CanRead
            // and an ordinary Response stream of a webserver can't do this
            // so we have to ask and provide a way around this
            if (stream.CanRead && stream.CanSeek && stream.CanWrite)
            {
                // all is fine the package can be created in a direct way
                CreatePackage(stream, true, _spreadsheetDocumentType, options);
            }
            else
            {
                // the harder way
                using var ms = new MemoryStream();
                CreatePackage(ms, true, _spreadsheetDocumentType, options);
                // not really necessary, because I changed CopyStream too.
                // but for better understanding and if somebody in the future
                // provide an changed version of CopyStream
                ms.Position = 0;
                CopyStream(ms, stream);
            }
        }
        else if (_loadSource == XlsxLoadSource.File)
        {
            using (var fileStream = new FileStream(_originalFile, FileMode.Open, FileAccess.Read))
            {
                CopyStream(fileStream, stream);
            }
            CreatePackage(stream, false, _spreadsheetDocumentType, options);
        }
        else if (_loadSource == XlsxLoadSource.Stream)
        {
            _originalStream.Position = 0;
            if (_originalStream != stream)
                CopyStream(_originalStream, stream);

            CreatePackage(stream, false, _spreadsheetDocumentType, options);
        }

        _loadSource = XlsxLoadSource.Stream;
        _originalStream = stream;
        _originalFile = null;
    }

    internal static void CopyStream(Stream input, Stream output)
    {
        var buffer = new byte[8 * 1024];
        int len;
        // dm 20130422, it is always a good idea to rewind the input stream, or not?
        if (input.CanSeek)
            input.Seek(0, SeekOrigin.Begin);
        while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
            output.Write(buffer, 0, len);
        // dm 20130422, and flushing the output after write
        output.Flush();
    }

    public IXlsxTable Table(string tableName, StringComparison comparisonType = StringComparison.OrdinalIgnoreCase)
    {
        ThrowIfDisposed();
        if (!TryGetTable(tableName, out XlsxTable table, comparisonType))
            throw new ArgumentOutOfRangeException($"Table {tableName} was not found.");

        return table;
    }

    /// <summary>
    /// Try to find a table with <paramref name="tableName"/> in a workbook.
    /// </summary>
    internal bool TryGetTable(string tableName, out XlsxTable table, StringComparison comparisonType = StringComparison.OrdinalIgnoreCase)
    {
        ThrowIfDisposed();
        table = WorksheetsInternal
            .SelectMany<XlsxWorksheet, XlsxTable>(ws => ws.Tables)
            .FirstOrDefault(t => t.Name.Equals(tableName, comparisonType));

        return table is not null;
    }

    /// <summary>
    /// Try to find a table that covers same area as the <paramref name="area"/> in a workbook.
    /// </summary>
    internal bool TryGetTable(SheetArea area, out XlsxTable foundTable)
    {
        ThrowIfDisposed();
        foreach (XlsxWorksheet sheet in WorksheetsInternal)
        {
            if (XlsxHelper.SheetComparer.Equals(sheet.Name, area.Name))
            {
                foreach (XlsxTable table in sheet.Tables)
                {
                    if (table.Area != area.Area)
                        continue;

                    foundTable = table;
                    return true;
                }

                // No other sheet has correct name.
                break;
            }
        }

        foundTable = null;
        return false;
    }

    public IXlsxWorksheet Worksheet(string name)
    {
        ThrowIfDisposed();
        return WorksheetsInternal.Worksheet(name);
    }

    public IXlsxWorksheet Worksheet(int position)
    {
        ThrowIfDisposed();
        return WorksheetsInternal.Worksheet(position);
    }

    public IXlsxCustomProperty CustomProperty(string name)
    {
        ThrowIfDisposed();
        return CustomProperties.CustomProperty(name);
    }

    public IXlsxCells FindCells(Func<IXlsxCell, bool> predicate)
    {
        ThrowIfDisposed();
        var cells = new XlsxCells(this, false, XlsxCellsUsedOptions.AllContents);
        foreach (XlsxWorksheet ws in WorksheetsInternal)
        {
            foreach (XlsxCell cell in ws.CellsUsed(XlsxCellsUsedOptions.All).Cast<XlsxCell>())
            {
                if (predicate(cell))
                    cells.Add(cell);
            }
        }
        return cells;
    }

    public IXlsxRows FindRows(Func<IXlsxRow, bool> predicate)
    {
        ThrowIfDisposed();
        var rows = new XlsxRows(this, worksheet: null, defaultStyleSheet: null);
        foreach (XlsxWorksheet ws in WorksheetsInternal)
        {
            foreach (IXlsxRow row in ws.Rows().Where(predicate))
                rows.Add(row as XlsxRow);
        }
        return rows;
    }

    public IXlsxColumns FindColumns(Func<IXlsxColumn, bool> predicate)
    {
        ThrowIfDisposed();
        var columns = new XlsxColumns(this, worksheet: null, defaultStyleSheet: null);
        foreach (XlsxWorksheet ws in WorksheetsInternal)
        {
            foreach (IXlsxColumn column in ws.Columns().Where(predicate))
                columns.Add(column as XlsxColumn);
        }
        return columns;
    }

    /// <summary>
    /// Searches the cells' contents for a given piece of text
    /// </summary>
    /// <param name="searchText">The search text.</param>
    /// <param name="compareOptions">The compare options.</param>
    /// <param name="searchFormulae">if set to <c>true</c> search formulae instead of cell values.</param>
    public IEnumerable<IXlsxCell> Search(string searchText, CompareOptions compareOptions = CompareOptions.Ordinal, bool searchFormulae = false)
    {
        ThrowIfDisposed();
        foreach (XlsxWorksheet ws in WorksheetsInternal)
        {
            foreach (IXlsxCell cell in ws.Search(searchText, compareOptions, searchFormulae))
                yield return cell;
        }
    }

    #region Fields

    private XlsxLoadSource _loadSource = XlsxLoadSource.New;
    private string _originalFile;
    private Stream _originalStream;
    private XlsxWorkbookProtection _workbookProtection;

    #endregion Fields

    #region Constructor

    /// <summary>
    ///   Creates a new Excel workbook.
    /// </summary>
    public XlsxWorkbook()
        : this(new LoadOptions())
    {
    }

    internal XlsxWorkbook(string file, bool asTemplate)
        : this(new LoadOptions())
    {
        Styles = new XlsxWorkbookStyles();
        LoadSheetsFromTemplate(file);
    }

    /// <summary>
    ///   Opens an existing workbook from a file.
    /// </summary>
    /// <param name = "file">The file to open.</param>
    public XlsxWorkbook(string file)
        : this(file, new LoadOptions())
    {
    }

    public XlsxWorkbook(string file, LoadOptions loadOptions)
        : this(loadOptions)
    {
        _loadSource = XlsxLoadSource.File;
        _originalFile = file;
        _spreadsheetDocumentType = GetWorkbookPackageType(_originalFile);
        Styles = new XlsxWorkbookStyles();
        Load(file);

        if (loadOptions.RecalculateAllFormulas)
            this.RecalculateAllFormulas();
    }

    /// <summary>
    ///   Opens an existing workbook from a stream.
    /// </summary>
    /// <param name = "stream">The stream to open.</param>
    public XlsxWorkbook(Stream stream)
        : this(stream, new LoadOptions())
    {
    }

    public XlsxWorkbook(Stream stream, LoadOptions loadOptions)
        : this(loadOptions)
    {
        _loadSource = XlsxLoadSource.Stream;
        _originalStream = stream;
        Styles = new XlsxWorkbookStyles();
        Load(stream);

        if (loadOptions.RecalculateAllFormulas)
            this.RecalculateAllFormulas();
    }

    public XlsxWorkbook(LoadOptions loadOptions)
    {
        ArgumentNullException.ThrowIfNull(loadOptions);

        DpiX = loadOptions.Dpi.X;
        DpiY = loadOptions.Dpi.Y;
        StrictAttributeParsing = loadOptions.StrictAttributeParsing;
        GraphicEngine = loadOptions.GraphicEngine ?? LoadOptions.DefaultGraphicEngine ?? DefaultGraphicEngine.Instance.Value;
        CancellationToken = loadOptions.CancellationToken;
        MaxCalculationMemoryBytes = loadOptions.MaxCalculationMemoryBytes;
        Protection = new XlsxWorkbookProtection(DefaultProtectionAlgorithm);
        DefaultRowHeight = 15;
        DefaultColumnWidth = 8.43;
        Styles = new XlsxWorkbookStyles();
        Styles.Initialize();
        RowHeight = DefaultRowHeight;
        ColumnWidth = DefaultColumnWidth;
        PageOptions = DefaultPageOptions;
        Outline = DefaultOutline;
        Properties = new XlsxWorkbookProperties();
        CalculateMode = XlsxCalculateMode.Default;
        ReferenceStyle = XlsxReferenceStyle.Default;
        InitializeTheme();
        ShowFormulas = DefaultShowFormulas;
        ShowGridLines = DefaultShowGridLines;
        ShowOutlineSymbols = DefaultShowOutlineSymbols;
        ShowRowColHeaders = DefaultShowRowColHeaders;
        ShowRuler = DefaultShowRuler;
        ShowWhiteSpace = DefaultShowWhiteSpace;
        ShowZeros = DefaultShowZeros;
        RightToLeft = DefaultRightToLeft;
        WorksheetsInternal = new XlsxWorksheets(this);
        DefinedNamesInternal = new XlsxDefinedNames(this);
        PivotCachesInternal = new XlsxPivotCaches(this);
        CustomProperties = new XlsxCustomProperties(this);
        ShapeIdManager = new XlsxIdManager();
        Author = Environment.UserName;
    }

    #endregion Constructor

    #region Nested type: UnsupportedSheet

    internal sealed class UnsupportedSheet
    {
        public bool IsActive;
        public uint SheetId;
        public int Position;
    }

    #endregion Nested type: UnsupportedSheet

    public IXlsxCell Cell(string namedCell)
    {
        ThrowIfDisposed();
        IXlsxDefinedName namedRange = DefinedName(namedCell);
        if (namedRange != null)
        {
            return namedRange.Ranges?.FirstOrDefault()?.FirstCell();
        }
        else
            return CellFromFullAddress(namedCell, out _);
    }

    public IXlsxCells Cells(string namedCells)
    {
        ThrowIfDisposed();
        return Ranges(namedCells).Cells();
    }

    public IXlsxRange Range(string range)
    {
        ThrowIfDisposed();
        IXlsxDefinedName namedRange = DefinedName(range);
        if (namedRange != null)
            return namedRange.Ranges.FirstOrDefault();
        else
            return RangeFromFullAddress(range, out _);
    }

    public IXlsxRanges Ranges(string ranges)
    {
        ThrowIfDisposed();
        var retVal = new XlsxRanges(this);
        var rangePairs = ranges.Split(',');
        foreach (IXlsxRange range in rangePairs.Select(r => Range(r.Trim())).Where(range => range != null))
        {
            retVal.Add(range);
        }
        return retVal;
    }

    internal XlsxIdManager ShapeIdManager
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        private set
        {
            ThrowIfDisposed();
            field = value;
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        foreach (XlsxWorksheet worksheet in WorksheetsInternal)
            worksheet.Cleanup();

        _disposed = true;
    }

    public bool Use1904DateSystem { get; set; }

    public XlsxWorkbook SetUse1904DateSystem()
    {
        return SetUse1904DateSystem(true);
    }

    public XlsxWorkbook SetUse1904DateSystem(bool value)
    {
        Use1904DateSystem = value;
        return this;
    }

    public IXlsxWorksheet AddWorksheet()
    {
        ThrowIfDisposed();
        return Worksheets.Add();
    }

    public IXlsxWorksheet AddWorksheet(int position)
    {
        ThrowIfDisposed();
        return Worksheets.Add(position);
    }

    public IXlsxWorksheet AddWorksheet(string sheetName)
    {
        ThrowIfDisposed();
        return Worksheets.Add(sheetName);
    }

    public IXlsxWorksheet AddWorksheet(string sheetName, int position)
    {
        ThrowIfDisposed();
        return Worksheets.Add(sheetName, position);
    }

    public void AddWorksheet(DataSet dataSet)
    {
        ThrowIfDisposed();
        Worksheets.Add(dataSet);
    }

    public void AddWorksheet(IXlsxWorksheet worksheet)
    {
        ThrowIfDisposed();
        worksheet.CopyTo(this, worksheet.Name);
    }

    public IXlsxWorksheet AddWorksheet(DataTable dataTable)
    {
        ThrowIfDisposed();
        return Worksheets.Add(dataTable);
    }

    public IXlsxWorksheet AddWorksheet(DataTable dataTable, string sheetName)
    {
        ThrowIfDisposed();
        return Worksheets.Add(dataTable, sheetName);
    }

    public IXlsxWorksheet AddWorksheet(DataTable dataTable, string sheetName, string tableName)
    {
        ThrowIfDisposed();
        return Worksheets.Add(dataTable, sheetName, tableName);
    }

    private XlsxCalcEngine _calcEngine;

    internal XlsxCalcEngine CalcEngine
    {
        get { return _calcEngine ??= new XlsxCalcEngine(CultureInfo.CurrentCulture); }
    }

    public XlsxCellValue Evaluate(string expression)
    {
        ThrowIfDisposed();
        return CalcEngine.EvaluateFormula(expression, this).ToCellValue();
    }

    /// <summary>
    /// Force recalculation of all cell formulas.
    /// </summary>
    public void RecalculateAllFormulas()
    {
        ThrowIfDisposed();
        foreach (XlsxWorksheet sheet in WorksheetsInternal)
            sheet.Internals.CellsCollection.FormulaSlice.MarkDirty(Area.Full);

        CalcEngine.Recalculate(this, null);
    }

    private static XlsxCalcEngine _calcEngineExpr;
    private WorkbookPackageType _spreadsheetDocumentType;

    private static XlsxCalcEngine CalcEngineExpr
    {
        get { return _calcEngineExpr ??= new XlsxCalcEngine(CultureInfo.InvariantCulture); }
    }

    /// <summary>
    /// Evaluate a formula and return a value. Formulas with references don't work and culture used for conversion is invariant.
    /// </summary>
    public static XlsxCellValue EvaluateExpr(string expression)
    {
        return CalcEngineExpr.EvaluateFormula(expression).ToCellValue();
    }

    /// <summary>
    /// Evaluate a formula and return a value. Use current culture.
    /// </summary>
    internal static XlsxCellValue EvaluateExprCurrent(string expression)
    {
        return new XlsxCalcEngine(CultureInfo.CurrentCulture).EvaluateFormula(expression).ToCellValue();
    }

    public string Author { get; set; }

    public bool LockStructure
    {
        get
        {
            ThrowIfDisposed();
            return Protection.IsProtected && !Protection.AllowedElements.HasFlag(XlsxWorkbookProtectionElements.Structure);
        }
        set
        {
            ThrowIfDisposed();
            if (!Protection.IsProtected)
                throw new InvalidOperationException($"Enable workbook protection before setting the {nameof(LockStructure)} property");

            Protection.AllowElement(XlsxWorkbookProtectionElements.Structure, value);
        }
    }

    public XlsxWorkbook SetLockStructure(bool value)
    {
        ThrowIfDisposed();
        LockStructure = value; return this;
    }

    public bool LockWindows
    {
        get
        {
            ThrowIfDisposed();
            return Protection.IsProtected && !Protection.AllowedElements.HasFlag(XlsxWorkbookProtectionElements.Windows);
        }
        set
        {
            ThrowIfDisposed();
            if (!Protection.IsProtected)
                throw new InvalidOperationException($"Enable workbook protection before setting the {nameof(LockWindows)} property");

            Protection.AllowElement(XlsxWorkbookProtectionElements.Windows, value);
        }
    }

    public XlsxWorkbook SetLockWindows(bool value)
    {
        ThrowIfDisposed();
        LockWindows = value;
        return this;
    }

    public bool IsPasswordProtected => Protection.IsPasswordProtected;
    public bool IsProtected => Protection.IsProtected;

    IXlsxWorkbookProtection IXlsxProtectable<IXlsxWorkbookProtection, XlsxWorkbookProtectionElements>.Protection
    {
        get
        {
            ThrowIfDisposed();
            return Protection;
        }
        set
        {
            ThrowIfDisposed();
            Protection = value as XlsxWorkbookProtection;
        }
    }

    internal XlsxWorkbookProtection Protection
    {
        get
        {
            ThrowIfDisposed();
            return _workbookProtection;
        }
        set
        {
            ThrowIfDisposed();
            _workbookProtection = value.Clone().CastTo<XlsxWorkbookProtection>();
        }
    }

    public IXlsxWorkbookProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm)
    {
        ThrowIfDisposed();
        return Protection.Protect(algorithm);
    }

    public IXlsxWorkbookProtection Protect(XlsxWorkbookProtectionElements allowedElements)
    {
        ThrowIfDisposed();
        return Protection.Protect(allowedElements);
    }

    public IXlsxWorkbookProtection Protect(Algorithm algorithm, XlsxWorkbookProtectionElements allowedElements)
    {
        ThrowIfDisposed();
        return Protection.Protect(algorithm, allowedElements);
    }

    public IXlsxWorkbookProtection Protect(string password, Algorithm algorithm = DefaultProtectionAlgorithm)
    {
        ThrowIfDisposed();
        return Protect(password, algorithm, XlsxWorkbookProtectionElements.Windows);
    }

    public IXlsxWorkbookProtection Protect(string password, Algorithm algorithm, XlsxWorkbookProtectionElements allowedElements)
    {
        ThrowIfDisposed();
        return Protection.Protect(password, algorithm, allowedElements);
    }

    IXlsxElementProtection IXlsxProtectable.Protect(Algorithm algorithm)
    {
        ThrowIfDisposed();
        return Protect(algorithm);
    }

    IXlsxElementProtection IXlsxProtectable.Protect(string password, Algorithm algorithm)
    {
        ThrowIfDisposed();
        return Protect(password, algorithm);
    }

    IXlsxWorkbookProtection IXlsxProtectable<IXlsxWorkbookProtection, XlsxWorkbookProtectionElements>.Protect(XlsxWorkbookProtectionElements allowedElements)
    {
        ThrowIfDisposed();
        return Protect(allowedElements);
    }

    IXlsxWorkbookProtection IXlsxProtectable<IXlsxWorkbookProtection, XlsxWorkbookProtectionElements>.Protect(Algorithm algorithm, XlsxWorkbookProtectionElements allowedElements)
    {
        ThrowIfDisposed();
        return Protect(algorithm, allowedElements);
    }

    IXlsxWorkbookProtection IXlsxProtectable<IXlsxWorkbookProtection, XlsxWorkbookProtectionElements>.Protect(string password, Algorithm algorithm, XlsxWorkbookProtectionElements allowedElements)
    {
        ThrowIfDisposed();
        return Protect(password, algorithm, allowedElements);
    }

    public IXlsxWorkbookProtection Unprotect()
    {
        ThrowIfDisposed();
        return Protection.Unprotect();
    }

    public IXlsxWorkbookProtection Unprotect(string password)
    {
        ThrowIfDisposed();
        return Protection.Unprotect(password);
    }

    IXlsxElementProtection IXlsxProtectable.Unprotect()
    {
        ThrowIfDisposed();
        return Unprotect();
    }

    IXlsxElementProtection IXlsxProtectable.Unprotect(string password)
    {
        ThrowIfDisposed();
        return Unprotect(password);
    }

    /// <summary>
    /// Notify various component of a workbook that sheet has been added.
    /// </summary>
    internal void NotifyWorksheetAdded(XlsxWorksheet newSheet)
    {
        ThrowIfDisposed();
        _calcEngine.OnAddedSheet(newSheet);
    }

    /// <summary>
    /// Notify various component of a workbook that sheet is about to be removed.
    /// </summary>
    internal void NotifyWorksheetDeleting(XlsxWorksheet sheet)
    {
        ThrowIfDisposed();
        _calcEngine.OnDeletingSheet(sheet);
    }

    public override string ToString()
    {
        ThrowIfDisposed();
        return _loadSource switch
        {
            XlsxLoadSource.New => "XlsxWorkbook(new)",
            XlsxLoadSource.File => string.Format("XlsxWorkbook({0})", _originalFile),
            XlsxLoadSource.Stream => string.Format("XlsxWorkbook({0})", _originalStream.ToString()),
            _ => throw new NotImplementedException(),
        };
    }

    internal XlsxCellFormat Format
    {
        get
        {
            ThrowIfDisposed();
            return XlsxCellFormat.ForWorkbook(this);
        }
    }

    private void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }
}
