#nullable disable

using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Calculation.Exceptions;

/// <summary>
/// Evaluation of the formula needs an information that wasn't available. That can happen if the formula
/// is evaluated from methods like <see cref="XlsxWorkbook.Evaluate(string)"/>. Causes vary, e.g. implicit intersection
/// needs an address of the formula cell. Various methods in Aid.Xlsx are missing different information, e.g.
/// <see cref="IXlsxWorksheet.Evaluate(string, string)"/> has worksheet, but no cell address (=ranges will work, other things won't).
/// </summary>
internal class MissingContextException : InvalidOperationException
{
}
