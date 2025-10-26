using System;
using System.Collections.Generic;
using System.Linq;
using SpreadsheetApp11.Core;

namespace SpreadsheetApp11.Expressions
{
    public class Evaluator
    {
        private readonly System.Func<Sheet> _sheetAccessor;
        private readonly HashSet<(int r, int c)> _stack = new();

        public Evaluator(System.Func<Sheet> sheetAccessor) { _sheetAccessor = sheetAccessor; }
        public double EvalCell(int r, int c)
        {
            try
            {
                var sheet = _sheetAccessor();
                if (sheet == null)
                    throw new EvalException("Sheet is null");

                var cell = sheet.GetCell(r, c);
                if (cell == null)
                    throw new EvalException("Cell is null");

                if (cell.Parsed == null && string.IsNullOrEmpty(cell.RawText))
                {
                    return 0;
                }

                if (cell.HasCachedValue)
                    return cell.CachedValue;

                if (!_stack.Add((r, c)))
                    throw new EvalException("Cycle detected");

                try
                {
                    double val = EvalExpr(cell.Parsed);

                    if (cell.Parsed != null || !string.IsNullOrEmpty(cell.RawText))
                    {
                        cell.CachedValue = val;
                        cell.HasCachedValue = true;
                    }

                    return val;
                }
                finally
                {
                    _stack.Remove((r, c));
                }
            }
            catch (EvalException ex)
            {
                if (ex.Message.Contains("Cycle"))
                {
                    throw;
                }
                return 0;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        private double EvalExpr(Expr expr)
        {
            if (expr == null) return 0;

            return expr switch
            {
                Literal l => l.Value,
                TextLiteral t => 0,
                CellRef cr => EvalCell(cr.Row, cr.Col),
                Range rg => SumRange(rg),
                Binary b => EvalBinary(b),
                FuncCall fc => EvalFuncCall(fc),
                _ => 0
            };
        }

        private double EvalBinary(Binary binary)
        {
            var left = EvalExpr(binary.Left);
            var right = EvalExpr(binary.Right);

            return binary.Op switch
            {
                '+' => left + right,
                '-' => left - right,
                '*' => left * right,
                '/' => right == 0 ? throw new EvalException("Division by zero") : left / right,
                '^' => Math.Pow(left, right),
                _ => throw new EvalException($"Unknown operator: {binary.Op}")
            };
        }

        private double EvalFuncCall(FuncCall fc)
        {
            var args = fc.Args.Select(EvalExpr).ToArray();

            return fc.Name.ToLower() switch
            {
                "min" when args.Length == 2 => Math.Min(args[0], args[1]),
                "max" when args.Length == 2 => Math.Max(args[0], args[1]),
                "inc" when args.Length == 1 => args[0] + 1,
                "dec" when args.Length == 1 => args[0] - 1,
                _ => throw new EvalException($"Unknown function: {fc.Name}")
            };
        }

        private IEnumerable<double> IterateRange(Range rg)
        {
            var r1 = (rg.From.Row < rg.To.Row) ? rg.From.Row : rg.To.Row;
            var r2 = (rg.From.Row > rg.To.Row) ? rg.From.Row : rg.To.Row;
            var c1 = (rg.From.Col < rg.To.Col) ? rg.From.Col : rg.To.Col;
            var c2 = (rg.From.Col > rg.To.Col) ? rg.From.Col : rg.To.Col;

            for (int r = r1; r <= r2; r++)
            {
                for (int c = c1; c <= c2; c++)
                {
                    var value = EvalCell(r, c);
                    if (value != 0)
                        yield return value;
                }
            }
        }

        private double SumRange(Range rg)
        {
            double sum = 0;
            foreach (var value in IterateRange(rg))
            {
                sum += value;
            }
            return sum;
        }
    }
}