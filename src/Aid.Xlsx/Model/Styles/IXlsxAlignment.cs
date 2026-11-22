namespace Aid.Xlsx.Model.Styles;

public enum XlsxAlignmentReadingOrderValues : byte
{
    ContextDependent,
    LeftToRight,
    RightToLeft
}

public enum XlsxAlignmentHorizontalValues : byte
{
    Center,
    CenterContinuous,

    /// <summary>
    /// <para>
    /// If <see cref="IXlsxAlignment.JustifyLastLine"/> is <c>false</c>, each line is justified
    /// (flushed left and right), regardless of whether line has explicit line break or not.
    /// This is a difference from <see cref="Justify"/> that only justifies automatically
    /// wrapped lines and doesn't justify last line.
    /// </para>
    /// <para>
    /// If <see cref="IXlsxAlignment.JustifyLastLine"/> is <c>true</c>, the words of each line
    /// are distributed along the line, so the gap on the left and right side of each word has
    /// same width. The gap is also on between word and left and right side of a cell.
    /// </para>
    /// </summary>
    /// <remarks>
    /// This alignment implicitly turns on <see cref="IXlsxAlignment.WrapText"/>, regardless of
    /// actual value.
    /// </remarks>
    Distributed,
    Fill,
    General,

    /// <summary>
    /// The text is justified when it is flushed left and right. For each line of text, it
    /// aligns each line of the wrapped text in a cell to the right and left (except the last
    /// line). Only text in wrapped lines (except last one) is justified. Text with explicit
    /// line breaks isn't.
    /// </summary>
    Justify,
    Left,
    Right
}

public enum XlsxAlignmentVerticalValues : byte
{
    Bottom,
    Center,
    Distributed,
    Justify,
    Top
}

public interface IXlsxAlignment : IEquatable<IXlsxAlignment>
{
    /// <summary>
    /// Gets or sets the cell's horizontal alignment.
    /// </summary>
    XlsxAlignmentHorizontalValues Horizontal { get; set; }

    /// <summary>
    /// Gets or sets the cell's vertical alignment.
    /// </summary>
    XlsxAlignmentVerticalValues Vertical { get; set; }

    /// <summary>
    /// Gets or sets the cell's text indentation. If <see cref="Horizontal"/>
    /// is <see cref="XlsxAlignmentHorizontalValues.General"/>, the value is
    /// changed to <see cref="XlsxAlignmentHorizontalValues.Left"/>.
    /// </summary>
    /// <value>
    /// The value means a width equal to <c>indent_value * 3 * width_of_space_in_normal_font</c>.
    /// </value>
    /// <exception cref="ArgumentOutOfRangeException">When the value is outside of [0,255].</exception>
    /// <exception cref="InvalidOperationException">If the value is positive and <see cref="Horizontal"/>
    ///     is either <see cref="XlsxAlignmentHorizontalValues.Left"/>, <see cref="XlsxAlignmentHorizontalValues.Right"/>
    ///     or <see cref="XlsxAlignmentHorizontalValues.Distributed"/>.
    /// </exception>
    Int32 Indent { get; set; }

    /// <summary>
    /// Gets or sets whether the cell's last line is justified or not.
    /// </summary>
    /// <value>
    /// The value changes the behavior of <see cref="XlsxAlignmentHorizontalValues.Distributed"/>
    /// alignment. Name of the property doesn't match actual behavior.
    /// </value>
    Boolean JustifyLastLine { get; set; }

    /// <summary>
    /// Gets or sets the cell's reading order.
    /// </summary>
    XlsxAlignmentReadingOrderValues ReadingOrder { get; set; }

    /// <summary>
    /// Gets or sets the cell's relative indent.
    /// </summary>
    /// <remarks>
    /// This property doesn't seem to work. It only set indent of a cell to 0 when used as part
    /// of differential format (regardless of relative ident value or actual ident value).
    /// </remarks>
    /// <value>
    /// The value is used only in differential formatting. It determines an additional indent
    /// to add to cells exiting <see cref="Indent"/>.
    /// </value>
    Int32 RelativeIndent { get; set; }

    /// <summary>
    /// Gets or sets whether the cell's font size should decrease to fit the contents.
    /// </summary>
    /// <remarks>
    /// Only applicable when <see cref="WrapText"/> is <c>false</c>, either directly or
    /// implicitly through some <see cref="XlsxAlignmentHorizontalValues"/> alignment (e.g.
    /// <see cref="XlsxAlignmentHorizontalValues.Distributed"/>).
    /// </remarks>
    Boolean ShrinkToFit { get; set; }

    /// <summary>
    /// Gets or sets the cell's text rotation in degrees.
    /// </summary>
    /// <value>
    /// Allowed values are -90 (text is rotated clockwise) to 90 (text is rotated
    /// counterclockwise) and 255 for vertical layout of a text.
    /// </value>
    /// <exception cref="ArgumentOutOfRangeException">When setting the rotation to a value outside of [-90,90] or 255.</exception>
    Int32 TextRotation { get; set; }

    /// <summary>
    /// Gets or sets whether the cell's text should wrap if it doesn't fit.
    /// </summary>
    Boolean WrapText { get; set; }

    /// <summary>
    /// Gets or sets whether the cell's text should be displayed from top to bottom
    /// (as opposed to the normal left to right).
    /// </summary>
    /// <remarks>
    /// The setter has same effect as as <c>alignment.TextRotation = topToBottom ? 255 : 0</c>.
    /// </remarks>
    Boolean TopToBottom { get; set; }

    /// <summary>
    /// Sets the cell's horizontal alignment.
    /// </summary>
    /// <inheritdoc cref="Horizontal"/>
    IXlsxStyle SetHorizontal(XlsxAlignmentHorizontalValues value);

    /// <summary>
    /// Sets the cell's vertical alignment.
    /// </summary>
    /// <inheritdoc cref="Vertical"/>
    IXlsxStyle SetVertical(XlsxAlignmentVerticalValues value);

    /// <summary>
    /// Sets the cell's text indentation.
    /// </summary>
    /// <inheritdoc cref="Indent"/>
    IXlsxStyle SetIndent(Int32 value);

    /// <summary>
    /// Changes mode of <see cref="XlsxAlignmentHorizontalValues.Distributed"/> alignment.
    /// </summary>
    /// <inheritdoc cref="JustifyLastLine"/>
    IXlsxStyle SetJustifyLastLine();

    /// <summary>
    /// Changes mode of <see cref="XlsxAlignmentHorizontalValues.Distributed"/> alignment.
    /// </summary>
    /// <inheritdoc cref="JustifyLastLine"/>
    IXlsxStyle SetJustifyLastLine(Boolean value);

    /// <summary>
    /// Sets the cell's reading order.
    /// </summary>
    /// <inheritdoc cref="ReadingOrder"/>
    IXlsxStyle SetReadingOrder(XlsxAlignmentReadingOrderValues value);

    IXlsxStyle SetRelativeIndent(Int32 value);

    /// <summary>
    /// Sets whether the cell's font size should decrease to fit the contents.
    /// </summary>
    /// <inheritdoc cref="ShrinkToFit"/>
    IXlsxStyle SetShrinkToFit();

    /// <summary>
    /// Sets whether the cell's font size should decrease to fit the contents.
    /// </summary>
    /// <inheritdoc cref="ShrinkToFit"/>
    IXlsxStyle SetShrinkToFit(Boolean value);

    /// <summary>
    /// Sets the cell's text rotation in degrees.
    /// </summary>
    /// <inheritdoc cref="TextRotation"/>
    IXlsxStyle SetTextRotation(Int32 value);

    /// <summary>
    /// Sets whether the cell's text should wrap if it doesn't fit.
    /// </summary>
    /// <inheritdoc cref="WrapText"/>
    IXlsxStyle SetWrapText();

    /// <summary>
    /// Sets whether the cell's text should wrap if it doesn't fit.
    /// </summary>
    /// <inheritdoc cref="WrapText"/>
    IXlsxStyle SetWrapText(Boolean value);

    /// <summary>
    /// Sets whether the cell's text should be displayed from top to bottom (as opposed to
    /// the normal left to right).
    /// </summary>
    /// <inheritdoc cref="TopToBottom"/>
    IXlsxStyle SetTopToBottom();

    /// <summary>
    /// Sets whether the cell's text should be displayed from top to bottom (as opposed to
    /// the normal left to right).
    /// </summary>
    /// <inheritdoc cref="TopToBottom"/>
    IXlsxStyle SetTopToBottom(Boolean value);
}
