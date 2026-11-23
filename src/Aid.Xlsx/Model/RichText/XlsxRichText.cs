using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.RichText;

/// <summary>
/// An API object for manipulating rich text. Every time it is changed, it calls
/// <see cref="OnContentChanged"/> to project changes back to the <see cref="SharedStringTable"/>.
/// </summary>
internal class XlsxRichText : XlsxFormattedText<IXlsxRichText>, IXlsxRichText
{
    // Should be set as the last thing in ctor to prevent firing changes to immutable rich text during ctor
    private readonly XlsxCell? _cell;

    /// <summary>
    /// Copy ctor to return user modifiable rich text from an immutable rich text stored
    /// in the shared string table.
    /// </summary>
    internal XlsxRichText(XlsxCell cell, XlsxFontFormatValue defaultFont, XlsxImmutableRichText original)
        : base(defaultFont, cell.Worksheet.Workbook.Styles)
    {
        foreach (XlsxImmutableRichText.RichTextRun originalRun in original.Runs)
        {
            var runText = original.GetRunText(originalRun);
            AddText(new XlsxRichString(runText, originalRun.Font, this, Styles, OnContentChanged));
        }

        var hasPhonetics = original.PhoneticRuns.Count > 0 || original.PhoneticsProperties.HasValue;
        if (hasPhonetics)
        {
            XlsxPhonetics phonetics;
            if (original.PhoneticsProperties.HasValue)
            {
                XlsxImmutableRichText.PhoneticProperties originalProps = original.PhoneticsProperties.Value;
                phonetics = new XlsxPhonetics(originalProps.Font, defaultFont, Styles, OnContentChanged)
                {
                    Type = originalProps.Type,
                    Alignment = originalProps.Alignment
                };
            }
            else
            {
                phonetics = new XlsxPhonetics(defaultFont, defaultFont, Styles, OnContentChanged);
            }

            foreach (XlsxImmutableRichText.PhoneticRun phoneticRun in original.PhoneticRuns)
                phonetics.Add(phoneticRun.Text, phoneticRun.StartIndex, phoneticRun.EndIndex);

            Phonetics = phonetics;
        }

        // TODO Styles: Convert to a factory method. The cell is set at the end to avoid false change trigger. Refactor so it's not needed anymore
        Container = this;
        _cell = cell;
    }

    internal XlsxRichText(XlsxCell cell, XlsxFontFormatValue defaultFont, string text)
        : this(cell, defaultFont)
    {
        AddText(new XlsxRichString(text, defaultFont, this, Styles, OnContentChanged));
    }

    internal XlsxRichText(XlsxCell cell, XlsxFontFormatValue defaultFont)
        : base(defaultFont, cell.Worksheet.Workbook.Styles)
    {
        Container = this;
        _cell = cell;
    }

    protected override void OnContentChanged()
    {
        // The rich text is still being created
        if (_cell is null)
            return;

        if (_cell.DataType != XlsxDataType.Text || !_cell.HasRichText)
            throw new InvalidOperationException("The rich text isn't a content of a cell.");

        _cell.SetOnlyValue(Text);
        Point point = _cell.Point;
        var richText = XlsxImmutableRichText.Create(this);
        _cell.Worksheet.Internals.CellsCollection.ValueSlice.SetRichText(point, richText);
    }
}
