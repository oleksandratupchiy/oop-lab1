using Antlr4.Runtime;
using SpreadsheetApp11.Expressions;
using SpreadsheetApp11.Core;
using System;

namespace SpreadsheetApp11.Parser
{
    public static class Calculator
    {
        public static double Evaluate(string expression, Sheet sheet)
        {
            try
            {
                var input = new AntlrInputStream(expression);
                var lexer = new MyGrammarLexer(input);
                var tokens = new CommonTokenStream(lexer);
                var parser = new MyGrammarParser(tokens);

                var tree = parser.start();
                var visitor = new MyGrammarVisitor(sheet);
                double result = visitor.Visit(tree);
                return result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Parser error: {ex.Message}");
            }
        }
    }

    public class MyGrammarVisitor : MyGrammarBaseVisitor<double>
    {
        private readonly Sheet _sheet;

        public MyGrammarVisitor(Sheet sheet)
        {
            _sheet = sheet;
        }

        public override double VisitStart(MyGrammarParser.StartContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitExpression(MyGrammarParser.ExpressionContext context)
        {
            return Visit(context.additiveExpression());
        }

        public override double VisitAdditiveExpression(MyGrammarParser.AdditiveExpressionContext context)
        {
            double result = Visit(context.multiplicativeExpression(0));

            for (int i = 1; i < context.multiplicativeExpression().Length; i++)
            {
                var op = context.GetChild(2 * i - 1).GetText();
                var right = Visit(context.multiplicativeExpression(i));

                if (op == "+")
                    result += right;
                else if (op == "-")
                    result -= right;
            }

            return result;
        }

        public override double VisitMultiplicativeExpression(MyGrammarParser.MultiplicativeExpressionContext context)
        {
            double result = Visit(context.powerExpression(0));

            for (int i = 1; i < context.powerExpression().Length; i++)
            {
                var op = context.GetChild(2 * i - 1).GetText();
                var right = Visit(context.powerExpression(i));

                if (op == "*")
                    result *= right;
                else if (op == "/")
                {
                    if (right == 0) throw new Exception("Division by zero");
                    result /= right;
                }
            }

            return result;
        }

        public override double VisitPowerExpression(MyGrammarParser.PowerExpressionContext context)
        {
            double result = Visit(context.unaryExpression(0));

            for (int i = 1; i < context.unaryExpression().Length; i++)
            {
                var right = Visit(context.unaryExpression(i));
                result = Math.Pow(result, right);
            }

            return result;
        }

        public override double VisitUnaryExpression(MyGrammarParser.UnaryExpressionContext context)
        {
            double result = Visit(context.primaryExpression());

            for (int i = 0; i < context.ChildCount - 1; i++)
            {
                var op = context.GetChild(i).GetText();
                if (op == "-")
                    result = -result;
            }

            return result;
        }

        public override double VisitPrimaryExpression(MyGrammarParser.PrimaryExpressionContext context)
        {
            if (context.NUMBER() != null)
            {
                var numberText = context.NUMBER().GetText();
                if (double.TryParse(numberText, out var value))
                    return value;
                return 0;
            }

            if (context.CELL_REF() != null)
            {
                var cellRef = context.CELL_REF().GetText();
                if (CellRefUtil.TryParse(cellRef, out var cr))
                {
                    var cell = _sheet.GetCell(cr.Row, cr.Col);
                    if (cell != null && cell.Parsed is Literal literal)
                    {
                        return literal.Value;
                    }
                    return 0;
                }
                throw new Exception($"Invalid cell reference: {cellRef}");
            }

            if (context.MIN() != null)
            {
                var arg1 = Visit(context.expression(0));
                var arg2 = Visit(context.expression(1));
                return Math.Min(arg1, arg2);
            }

            if (context.MAX() != null)
            {
                var arg1 = Visit(context.expression(0));
                var arg2 = Visit(context.expression(1));
                return Math.Max(arg1, arg2);
            }

            if (context.INC() != null)
            {
                var arg = Visit(context.expression(0));
                return arg + 1;
            }

            if (context.DEC() != null)
            {
                var arg = Visit(context.expression(0));
                return arg - 1;
            }

            if (context.expression() != null && context.expression().Length == 1)
            {
                return Visit(context.expression(0));
            }

            throw new Exception($"Unsupported primary expression: {context.GetText()}");
        }
    }
}