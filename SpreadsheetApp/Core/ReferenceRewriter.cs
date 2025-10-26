using System;
using System.Linq;
using SpreadsheetApp11.Expressions;

namespace SpreadsheetApp11.Core
{
    public static class ReferenceRewriter
    {
        public static SpreadsheetApp11.Expressions.Expr ShiftRows(SpreadsheetApp11.Expressions.Expr expr, int atIndex, int delta)
        {
            return Visit(expr, rDelta: delta, cDelta: 0, rowStart: atIndex, colStart: int.MaxValue);
        }

        public static SpreadsheetApp11.Expressions.Expr ShiftCols(SpreadsheetApp11.Expressions.Expr expr, int atIndex, int delta)
        {
            return Visit(expr, rDelta: 0, cDelta: delta, rowStart: int.MaxValue, colStart: atIndex);
        }

        private static SpreadsheetApp11.Expressions.Expr Visit(SpreadsheetApp11.Expressions.Expr e, int rDelta, int cDelta, int rowStart, int colStart)
        {

            return e switch
            {
                SpreadsheetApp11.Expressions.Literal l => l,
                SpreadsheetApp11.Expressions.TextLiteral t => t,
                SpreadsheetApp11.Expressions.CellRef cr => ShiftCellRef(cr, rDelta, cDelta, rowStart, colStart),
                SpreadsheetApp11.Expressions.Range rg => ShiftRange(rg, rDelta, cDelta, rowStart, colStart),
                SpreadsheetApp11.Expressions.Unary u => new SpreadsheetApp11.Expressions.Unary(u.Op, Visit(u.Arg, rDelta, cDelta, rowStart, colStart)),
                SpreadsheetApp11.Expressions.Binary b => new SpreadsheetApp11.Expressions.Binary(b.Op,
                    Visit(b.Left, rDelta, cDelta, rowStart, colStart),
                    Visit(b.Right, rDelta, cDelta, rowStart, colStart)),
                SpreadsheetApp11.Expressions.FuncCall fc => new SpreadsheetApp11.Expressions.FuncCall(fc.Name,
                    fc.Args.Select(arg => Visit(arg, rDelta, cDelta, rowStart, colStart)).ToArray()),
                _ => e
            };
        }

        private static SpreadsheetApp11.Expressions.CellRef ShiftCellRef(SpreadsheetApp11.Expressions.CellRef cr, int rDelta, int cDelta, int rowStart, int colStart)
        {
            var nr = cr.Row + (cr.Row >= rowStart ? rDelta : 0);
            var nc = cr.Col + (cr.Col >= colStart ? cDelta : 0);
            return new SpreadsheetApp11.Expressions.CellRef(nr, nc);
        }

        private static SpreadsheetApp11.Expressions.Range ShiftRange(SpreadsheetApp11.Expressions.Range rg, int rDelta, int cDelta, int rowStart, int colStart)
        {
            var from = ShiftCellRef(rg.From, rDelta, cDelta, rowStart, colStart);
            var to = ShiftCellRef(rg.To, rDelta, cDelta, rowStart, colStart);
            return new SpreadsheetApp11.Expressions.Range(from, to);
        }
    }
}