using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.DataValidation;

internal class XlsxDataValidation : IXlsxDataValidation
{
    private readonly XlsxWorksheet _worksheet;

    internal XlsxDataValidation(XlsxWorksheet worksheet)
    {
        _worksheet = worksheet;
        Initialize();
    }

    internal XlsxWorksheet Worksheet => _worksheet;

    internal XlsxAreaList Areas { get; set; } = XlsxAreaList.Empty;

    public void Clear()
    {
        Initialize();
    }

    internal void CopyFrom(IXlsxDataValidation dataValidation)
    {
        if (dataValidation == this) return;

        if (Areas.Count == 0)
            AddRanges(dataValidation.Ranges);

        IgnoreBlanks = dataValidation.IgnoreBlanks;
        InCellDropdown = dataValidation.InCellDropdown;
        ShowErrorMessage = dataValidation.ShowErrorMessage;
        ShowInputMessage = dataValidation.ShowInputMessage;
        InputTitle = dataValidation.InputTitle;
        InputMessage = dataValidation.InputMessage;
        ErrorTitle = dataValidation.ErrorTitle;
        ErrorMessage = dataValidation.ErrorMessage;
        ErrorStyle = dataValidation.ErrorStyle;
        AllowedValues = dataValidation.AllowedValues;
        Operator = dataValidation.Operator;
        MinValue = dataValidation.MinValue;
        MaxValue = dataValidation.MaxValue;
    }

    public bool IsDirty()
    {
        return
            AllowedValues != XlsxAllowedValues.AnyValue
            || (ShowInputMessage &&
               (!string.IsNullOrWhiteSpace(InputTitle) || !string.IsNullOrWhiteSpace(InputMessage)))
            || (ShowErrorMessage &&
               (!string.IsNullOrWhiteSpace(ErrorTitle) || !string.IsNullOrWhiteSpace(ErrorMessage)));
    }

    [MemberNotNull(nameof(minValue), nameof(maxValue), nameof(InputTitle), nameof(InputMessage), nameof(ErrorTitle), nameof(ErrorMessage))]
    private void Initialize()
    {
        AllowedValues = XlsxAllowedValues.AnyValue;
        IgnoreBlanks = true;
        ShowErrorMessage = true;
        ShowInputMessage = true;
        InCellDropdown = true;
        InputTitle = string.Empty;
        InputMessage = string.Empty;
        ErrorTitle = string.Empty;
        ErrorMessage = string.Empty;
        ErrorStyle = XlsxErrorStyle.Stop;
        Operator = XlsxOperator.Between;
        Value = string.Empty;
        minValue = string.Empty;
        maxValue = string.Empty;
    }

    #region IXlsxDataValidation Members

    private string maxValue;
    private string minValue;
    public XlsxAllowedValues AllowedValues { get; set; }

    public XlsxDateCriteria Date
    {
        get
        {
            AllowedValues = XlsxAllowedValues.Date;
            return new XlsxDateCriteria(this);
        }
    }

    public XlsxDecimalCriteria Decimal
    {
        get
        {
            AllowedValues = XlsxAllowedValues.Decimal;
            return new XlsxDecimalCriteria(this);
        }
    }

    public string ErrorMessage { get; set; }
    public XlsxErrorStyle ErrorStyle { get; set; }
    public string ErrorTitle { get; set; }
    public bool IgnoreBlanks { get; set; }
    public bool InCellDropdown { get; set; }
    public string InputMessage { get; set; }
    public string InputTitle { get; set; }
    public string MaxValue { get => maxValue; set { Validate(value); maxValue = value; } }
    public string MinValue { get => minValue; set { Validate(value); minValue = value; } }
    public XlsxOperator Operator { get; set; }

    public IEnumerable<IXlsxRange> Ranges
    {
        get
        {
            var ranges = new XlsxRanges(Worksheet);
            foreach (Area area in Areas)
            {
                XlsxRange range = _worksheet.Range(XlsxRangeAddress.FromSheetRange(_worksheet, area));
                ranges.Add(range);
            }

            return ranges;
        }
    }

    public bool ShowErrorMessage { get; set; }

    public bool ShowInputMessage { get; set; }

    public XlsxTextLengthCriteria TextLength
    {
        get
        {
            AllowedValues = XlsxAllowedValues.TextLength;
            return new XlsxTextLengthCriteria(this);
        }
    }

    public XlsxTimeCriteria Time
    {
        get
        {
            AllowedValues = XlsxAllowedValues.Time;
            return new XlsxTimeCriteria(this);
        }
    }

    public string Value
    {
        get { return MinValue; }
        set { MinValue = value; }
    }

    public XlsxWholeNumberCriteria WholeNumber
    {
        get
        {
            AllowedValues = XlsxAllowedValues.WholeNumber;
            return new XlsxWholeNumberCriteria(this);
        }
    }

    /// <summary>
    /// Add a range to the collection of ranges this rule applies to.
    /// If the specified range does not belong to the worksheet of the data validation
    /// rule it is transferred to the target worksheet.
    /// </summary>
    /// <param name="range">A range to add.</param>
    public void AddRange(IXlsxRange range)
    {
        ArgumentNullException.ThrowIfNull(range);

        // Do not add area if the DV has been detached (e.g. consolidation).
        var isDetached = !_worksheet.DataValidations.Contains(this);
        if (isDetached)
            return;

        // Ignore sheet of a range
        var area = Area.FromRangeAddress(range.RangeAddress);
        _worksheet.DataValidations.AddArea(this, area);
    }

    /// <summary>
    /// Add a collection of ranges to the collection of ranges this rule applies to.
    /// Ranges that do not belong to the worksheet of the data validation
    /// rule are transferred to the target worksheet.
    /// </summary>
    /// <param name="ranges">Ranges to add.</param>
    public void AddRanges(IEnumerable<IXlsxRange> ranges)
    {
        ranges ??= [];

        foreach (IXlsxRange range in ranges)
        {
            AddRange(range);
        }
    }

    /// <summary>
    /// Detach data validation rule of all ranges it applies to.
    /// </summary>
    public void ClearRanges()
    {
        Areas = XlsxAreaList.Empty;
    }

    public void Custom(string customValidation)
    {
        AllowedValues = XlsxAllowedValues.Custom;
        Value = customValidation;
    }

    public void List(string list)
    {
        List(list, true);
    }

    public void List(string list, bool inCellDropdown)
    {
        AllowedValues = XlsxAllowedValues.List;
        InCellDropdown = inCellDropdown;
        Value = list;
    }

    public void List(IXlsxRange range)
    {
        List(range, true);
    }

    public void List(IXlsxRange range, bool inCellDropdown)
    {
        List(range.RangeAddress.ToStringFixed(XlsxReferenceStyle.A1, true));
    }

    /// <summary>
    /// Remove the specified range from the collection of range this rule applies to.
    /// </summary>
    /// <param name="range">A range to remove.</param>
    public bool RemoveRange(IXlsxRange range)
    {
        if (range == null)
            return false;

        Area areaToDelete = SheetArea.From(range).Area;
        XlsxAreaList originalAreas = Areas;
        Areas = Areas.Without(areaToDelete);
        var deleted = originalAreas.Count > Areas.Count;
        return deleted;
    }

    #endregion IXlsxDataValidation Members

    private static void Validate(string value)
    {
        if (value.Length > 255)
            throw new ArgumentOutOfRangeException(nameof(value), "The maximum allowed length of the value is 255 characters.");
    }
}
