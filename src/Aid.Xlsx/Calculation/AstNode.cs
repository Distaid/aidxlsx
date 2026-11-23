using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.Formulas.References;
using Aid.Xlsx.Model.DefinedNames;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Calculation;

/// <summary>
/// Base class for all AST nodes. All AST nodes must be immutable.
/// </summary>
internal abstract class AstNode
{
    /// <summary>
    /// Method to accept a visitor (=call a method of visitor with correct type of the node).
    /// </summary>
    public abstract TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor);
}

/// <summary>
/// A base class for all AST nodes that can be evaluated to produce a value.
/// </summary>
internal abstract class ValueNode : AstNode
{
}

/// <summary>
/// AST node that contains a blank, logical, number, text or an error value.
/// </summary>
internal class ScalarNode(ScalarValue value) : ValueNode
{
    public ScalarValue Value { get; } = value;

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}

/// <summary>
/// AST node that contains a constant array. Array is at least 1x1.
/// </summary>
internal class ArrayNode(Array value) : ValueNode
{
    public Array Value { get; } = value;

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}

internal enum UnaryOp
{
    Add,
    Subtract,
    Percentage,
    SpillRange,
    ImplicitIntersection
}

/// <summary>
/// Unary expression, e.g. +123
/// </summary>
internal class UnaryNode(UnaryOp operation, ValueNode expr) : ValueNode
{
    public UnaryOp Operation { get; } = operation;

    public ValueNode Expression { get; } = expr;

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}

internal enum BinaryOp
{
    // Text operators
    Concat,
    // Arithmetic
    Add,
    Sub,
    Mult,
    Div,
    Exp,
    // Comparison operators
    Lt,
    Lte,
    Eq,
    Neq,
    Gte,
    Gt,
    // References operators
    Range,
    Union,
    Intersection
}

/// <summary>
/// Binary expression, e.g. 1+2
/// </summary>
internal class BinaryNode(BinaryOp operation, ValueNode exprLeft, ValueNode exprRight) : ValueNode
{
    public BinaryOp Operation { get; } = operation;

    public ValueNode LeftExpression { get; } = exprLeft;

    public ValueNode RightExpression { get; } = exprRight;

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}

/// <summary>
/// A function call, e.g. <c>SIN(0.5)</c>.
/// </summary>
internal class FunctionNode(PrefixNode? prefix, string name, IReadOnlyList<ValueNode> parms) : ValueNode
{
    public FunctionNode(string name, IReadOnlyList<ValueNode> parms) : this(null, name, parms)
    {
    }

    public PrefixNode? Prefix { get; } = prefix;

    /// <summary>
    /// Name of the function.
    /// </summary>
    public string Name { get; } = name;

    /// <summary>
    /// AST nodes for arguments of the function.
    /// </summary>
    public IReadOnlyList<ValueNode> Parameters { get; } = parms;

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}

/// <summary>
/// An placeholder node for AST nodes that are not yet supported in Aid.Xlsx.
/// </summary>
internal class NotSupportedNode(string featureName) : ValueNode
{
    public string FeatureName { get; } = featureName;

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}

/// <summary>
/// AST node for an reference to an external file in a formula.
/// </summary>
internal class FileNode : AstNode
{
    /// <summary>
    /// If the file is references indirectly, numeric identifier of a file.
    /// </summary>
    public int? Numeric { get; }

    /// <summary>
    /// If a file is referenced directly, a path to the file on the disc/UNC/web link, .
    /// </summary>
    public string? Path { get; }

    public FileNode(string path)
    {
        Path = path;
    }

    public FileNode(int numeric)
    {
        Numeric = numeric;
    }

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}

/// <summary>
/// AST node for prefix of a reference in a formula. Prefix is a specification where to look for a reference.
/// <list type="bullet">
/// <item>Prefix specifies a <c>Sheet</c> - used for references in the local workbook.</item>
/// <item>Prefix specifies a <c>FirstSheet</c> and a <c>LastSheet</c> - 3D reference, references uses all sheets between first and last.</item>
/// <item>Prefix specifies a <c>File</c>, no sheet is specified - used for named ranges in external file.</item>
/// <item>Prefix specifies a <c>File</c> and a <c>Sheet</c> - references looks for its address in the sheet of the file.</item>
/// </list>
/// </summary>
internal class PrefixNode(FileNode? file, string? sheet, string? firstSheet, string? lastSheet) : AstNode
{

    /// <summary>
    /// If prefix references data from another file, can be empty.
    /// </summary>
    public FileNode? File { get; } = file;

    /// <summary>
    /// Name of the sheet, without ! or escaped quotes. Can be null in some cases e.g. reference to a named range in an another file).
    /// </summary>
    public string? Sheet { get; } = sheet;

    /// <summary>
    /// If the prefix is for 3D reference, name of first sheet. Empty otherwise.
    /// </summary>
    public string? FirstSheet { get; } = firstSheet;

    /// <summary>
    /// If the prefix is for 3D reference, name of the last sheet. Empty otherwise.
    /// </summary>
    public string? LastSheet { get; } = lastSheet;

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);

    internal OneOf<IXlsxWorksheet, XlsxError> GetWorksheet(XlsxWorkbook wb)
    {
        if (File is not null)
            throw new NotImplementedException("References from other files are not yet implemented.");

        if (FirstSheet is not null || LastSheet is not null)
            throw new NotImplementedException("3D references are not yet implemented.");

        if (!wb.TryGetWorksheet(Sheet, out XlsxWorksheet worksheet))
            return XlsxError.CellReference;

        return OneOf<IXlsxWorksheet, XlsxError>.FromT0(worksheet);
    }
}

/// <summary>
/// AST node for a reference of an area in some sheet.
/// </summary>
internal class ReferenceNode(PrefixNode? prefix, ReferenceArea referenceArea, bool isA1) : ValueNode
{

    /// <summary>
    /// An optional prefix for reference item.
    /// </summary>
    public PrefixNode? Prefix { get; } = prefix;

    /// <summary>
    /// An address of a reference that corresponds to <see cref="Type"/>. Always without sheet (that is in the prefix).
    /// </summary>
    public string Address { get; } = isA1 ? referenceArea.GetDisplayStringA1() : referenceArea.GetDisplayStringR1C1();

    /// <summary>
    /// An area from a parser.
    /// </summary>
    public ReferenceArea ReferenceArea { get; } = referenceArea;

    /// <summary>
    /// Is the reference in A1 style? If <c>false</c>, then it is R1C1.
    /// </summary>
    public bool IsA1 { get; } = isA1;

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);

    public AnyValue GetReference(CalcContext ctx)
    {
        if (Prefix is null)
            return new Reference(new XlsxRangeAddress(null, Address));

        if (!Prefix.GetWorksheet(ctx.Workbook).TryPickT0(out var ws, out var err))
            return err;

        // TODO: XlsxRangeAddress can parse all types of reference item type, utilize known type for faster parsing + cache
        return new Reference(new XlsxRangeAddress((XlsxWorksheet)ws!, Address));
    }
}

/// <summary>
/// A name node in the formula. Name can refers to a generic formula, in most cases a reference, but it can be any kind of calculation (e.g. <c>A1+7</c>).
/// </summary>
internal class NameNode(PrefixNode? prefix, string name) : ValueNode
{

    /// <summary>
    /// An optional prefix for reference item.
    /// </summary>
    public PrefixNode? Prefix { get; } = prefix;

    public string Name { get; } = name;

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);

    public AnyValue GetValue(XlsxWorksheet ctxWs, XlsxCalcEngine engine)
    {
        var worksheet = ctxWs;
        if (Prefix is not null)
        {
            if (!Prefix.GetWorksheet(ctxWs.Workbook).TryPickT0(out var ws, out var err))
                return err;

            worksheet = (XlsxWorksheet)ws!;
        }

        if (!TryGetNameRange(worksheet, out var definedName))
            return XlsxError.NameNotRecognized;

        // Parser needs an equal sign for a union of ranges (or braces around formula)
        var nameFormula = definedName.RefersTo;
        nameFormula = nameFormula.StartsWith("=") ? nameFormula : "=" + nameFormula;
        return engine.EvaluateName(nameFormula, ctxWs);
    }

    internal bool TryGetNameRange(IXlsxWorksheet ws, [NotNullWhen(true)] out IXlsxDefinedName? definedName)
    {
        if (ws.DefinedNames.TryGetValue(Name, out var sheetDefinedName))
        {
            definedName = sheetDefinedName;
            return true;
        }

        if (ws.Workbook.DefinedNamesInternal.TryGetValue(Name, out var bookDefinedName))
        {
            definedName = bookDefinedName;
            return true;
        }

        definedName = null;
        return false;
    }
}

internal class StructuredReferenceNode(PrefixNode? prefix, string? table, StructuredReferenceArea area, string? firstColumn, string? lastColumn) : ValueNode
{

    /// <summary>
    /// Can be empty if no prefix available.
    /// </summary>
    public PrefixNode? Prefix { get; } = prefix;

    /// <summary>
    /// Table of the reference. It can be empty, if formula using the reference is within
    /// the table itself (e.g. total formulas).
    /// </summary>
    public string? Table { get; } = table;

    /// <summary>
    /// Area of the table that is considered for the range of cell of reference.
    /// </summary>
    public StructuredReferenceArea Area { get; } = area;

    /// <summary>
    /// First column of column range. If the reference refers to the whole table,
    /// the value is null.
    /// </summary>
    public string? FirstColumn { get; } = firstColumn;

    /// <summary>
    /// Last column of column range. If structured reference refers only to one column,
    /// it is same as <see cref="FirstColumn"/>. If the reference refers to the whole table,
    /// the value is null.
    /// </summary>
    public string? LastColumn { get; } = lastColumn;

    public override TResult Accept<TContext, TResult>(TContext context, IFormulaVisitor<TContext, TResult> visitor) => visitor.Visit(context, this);
}
