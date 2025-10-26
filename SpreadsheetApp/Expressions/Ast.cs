using System;
using System.Linq;

namespace SpreadsheetApp11.Expressions
{
    public abstract record Expr;

    public sealed record Literal(double Value) : Expr;
    public sealed record TextLiteral(string Text) : Expr;
    public sealed record CellRef(int Row, int Col) : Expr;
    public sealed record Range(CellRef From, CellRef To) : Expr;
    public sealed record Unary(char Op, Expr Arg) : Expr;
    public sealed record Binary(char Op, Expr Left, Expr Right) : Expr;
    public sealed record FuncCall(string Name, Expr[] Args) : Expr;

    public static class CellRefUtil
    {
        public static bool TryParse(string ident, out CellRef cr)
        {
            cr = default;
            int i = 0;
            while (i < ident.Length && char.IsLetter(ident[i])) i++;
            if (i == 0 || i == ident.Length) return false;

            string letters = ident[..i].ToUpperInvariant();
            string digits = ident[i..];
            if (!int.TryParse(digits, out int oneBasedRow) || oneBasedRow <= 0) return false;

            int col = 0;
            foreach (var ch in letters) col = col * 26 + (ch - 'A' + 1);
            cr = new CellRef(oneBasedRow - 1, col - 1);
            return true;
        }

        public static string ToA1(CellRef r)
        {
            int c = r.Col + 1;
            string name = "";
            while (c > 0)
            {
                int rem = (c - 1) % 26;
                name = (char)('A' + rem) + name;
                c = (c - 1) / 26;
            }
            return name + (r.Row + 1);
        }
    }

    public class ParseException : Exception { public ParseException(string m) : base(m) { } }
    public class EvalException : Exception { public EvalException(string m) : base(m) { } }

    public static class ExprFormatter
    {
        public static string Format(Expr expr)
        {
            return expr switch
            {
                Literal l => l.Value.ToString(),
                TextLiteral t => t.Text,
                CellRef cr => CellRefUtil.ToA1(cr),
                Range r => $"{CellRefUtil.ToA1(r.From)}:{CellRefUtil.ToA1(r.To)}",
                Binary b => $"({Format(b.Left)}{b.Op}{Format(b.Right)})",
                Unary u => $"{u.Op}{Format(u.Arg)}",
                FuncCall f => $"{f.Name}({string.Join(",", f.Args.Select(Format))})",
                _ => "#ERR"
            };
        }
    }
}