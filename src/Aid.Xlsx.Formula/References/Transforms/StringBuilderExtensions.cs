using System;
using System.Collections.Generic;
using System.Text;

namespace Aid.Xlsx.Formulas.References.Transforms;

/// <summary>
/// Extension methods for building formulas.
/// </summary>
internal static class StringBuilderExtensions
{
    extension(StringBuilder sb)
    {
        public StringBuilder AppendSheetReference(string? sheetName)
        {
            if (sheetName is null)
                return sb.Append("#REF!");

            return NameUtils.EscapeName(sb, sheetName).AppendReferenceSeparator();
        }

        public StringBuilder AppendExternalSheetReference(int workbookIndex, string sheetName)
        {
            if (NameUtils.ShouldQuote(sheetName.AsSpan()))
            {
                return sb
                    .Append('\'')
                    .AppendBookIndex(workbookIndex)
                    .AppendEscapedSheetName(sheetName)
                    .Append('\'')
                    .AppendReferenceSeparator();
            }

            return sb
                .AppendBookIndex(workbookIndex)
                .AppendSheetReference(sheetName);
        }

        public StringBuilder AppendEscapedSheetName(string sheetName)
        {
            var startIndex = sb.Length;
            return sb.Append(sheetName).Replace("'", "''", startIndex, sheetName.Length);
        }

        public StringBuilder AppendReferenceSeparator()
        {
            return sb.Append('!');
        }

        public StringBuilder AppendBookIndex(int bookIndex)
        {
            return sb.Append('[').Append(bookIndex).Append(']');
        }

        public StringBuilder AppendFunction(ModContext ctx, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<TransformedSymbol> arguments)
        {
            return sb.Append(functionName).AppendArguments(ctx, range, arguments);
        }

        public StringBuilder AppendRef(ReferenceArea? reference)
        {
            return reference is null ? sb.Append("#REF!") : reference.Value.Append(sb);
        }

        public StringBuilder AppendRef(RowCol? rowCol)
        {
            if (rowCol is null)
                return sb.Append("#REF!");

            rowCol.Value.Append(sb);
            return sb;
        }

        public StringBuilder AppendStartFragment(ModContext ctx, SymbolRange symbolRange, TransformedSymbol nestedNode)
        {
            return sb.AppendFragment(ctx.Formula, symbolRange.Start, nestedNode.OriginalRange.Start);
        }

        public StringBuilder AppendMiddleFragment(ModContext ctx, TransformedSymbol beforeNode, TransformedSymbol afterNode)
        {
            return sb.AppendFragment(ctx.Formula, beforeNode.OriginalRange.End, afterNode.OriginalRange.Start);
        }

        public StringBuilder AppendEndFragment(ModContext ctx, SymbolRange symbolRange, TransformedSymbol nestedNode)
        {
            return sb.AppendFragment(ctx.Formula, nestedNode.OriginalRange.End, symbolRange.End);
        }

        public StringBuilder AppendArguments(ModContext ctx, SymbolRange range, IReadOnlyList<TransformedSymbol> arguments)
        {
            if (arguments.Count == 0)
            {
                var braceIdx = GetStartBraceIndex(ctx, range, range.End);
                sb.AppendFragment(ctx.Formula, braceIdx, range.End);
            }
            else
            {
                sb
                    .AppendStartBrace(ctx, range, arguments[0])
                    .AppendArguments(ctx, arguments)
                    .AppendEndFragment(ctx, range, arguments[arguments.Count - 1]);
            }

            return sb;
        }

        private StringBuilder AppendStartBrace(ModContext ctx, SymbolRange range, TransformedSymbol firstNode)
        {
            var firstNodeStart = firstNode.OriginalRange.Start;
            var braceIdx = GetStartBraceIndex(ctx, range, firstNodeStart);
            return sb.AppendFragment(ctx.Formula, braceIdx, firstNodeStart);
        }

        private StringBuilder AppendFragment(string formula, int start, int end)
        {
            if (end > start)
                sb.Append(formula.AsSpan(start, end - start));

            return sb;
        }

        private StringBuilder AppendArguments(ModContext ctx, IReadOnlyList<TransformedSymbol> arguments)
        {
            if (arguments.Count > 0)
                sb.Append(arguments[0].AsSpan());

            for (var i = 1; i < arguments.Count; ++i)
            {
                sb.AppendMiddleFragment(ctx, arguments[i - 1], arguments[i]);
                sb.Append(arguments[i].AsSpan());
            }

            return sb;
        }
    }

    private static int GetStartBraceIndex(ModContext ctx, SymbolRange range, int nodeStart)
    {
        var formula = ctx.Formula;
        var braceIdx = nodeStart - 1;
        for (; braceIdx > range.Start; --braceIdx)
        {
            if (formula[braceIdx] == '(')
                return braceIdx;
        }

        throw new InvalidOperationException("No opening brace found.");
    }
}
