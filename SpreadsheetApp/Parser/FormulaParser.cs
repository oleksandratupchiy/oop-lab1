using Antlr4.Runtime;
using SpreadsheetApp11.Expressions;
using SpreadsheetApp11.Core;
using System;

namespace SpreadsheetApp11.Parser
{
    public static class FormulaParser
    {
        public static Expr Parse(string expression, Sheet sheet)
        {
            try
            {
                var input = new AntlrInputStream(expression);
                var lexer = new MyGrammarLexer(input);
                var tokens = new CommonTokenStream(lexer);
                var parser = new MyGrammarParser(tokens);

                var tree = parser.start();
                var visitor = new FormulaASTVisitor(sheet);
                return visitor.Visit(tree);
            }
            catch (Exception ex)
            {
                throw new Exception($"Parser error: {ex.Message}");
            }
        }
    }

    public class FormulaASTVisitor : MyGrammarBaseVisitor<Expr>
    {
        private readonly Sheet _sheet;

        public FormulaASTVisitor(Sheet sheet)
        {
            _sheet = sheet;
        }

        public override Expr VisitStart(MyGrammarParser.StartContext context)
        {
            return Visit(context.expression());
        }

        public override Expr VisitExpression(MyGrammarParser.ExpressionContext context)
        {
            return Visit(context.additiveExpression());
        }

        public override Expr VisitAdditiveExpression(MyGrammarParser.AdditiveExpressionContext context)
        {
            if (context.multiplicativeExpression().Length == 1)
            {
                return Visit(context.multiplicativeExpression(0));
            }

            Expr left = Visit(context.multiplicativeExpression(0));

            for (int i = 1; i < context.multiplicativeExpression().Length; i++)
            {
                var op = context.GetChild(2 * i - 1).GetText();
                var right = Visit(context.multiplicativeExpression(i));

                left = new Binary(op[0], left, right);
            }

            return left;
        }

        public override Expr VisitMultiplicativeExpression(MyGrammarParser.MultiplicativeExpressionContext context)
        {
            if (context.powerExpression().Length == 1)
            {
                return Visit(context.powerExpression(0));
            }

            Expr left = Visit(context.powerExpression(0));

            for (int i = 1; i < context.powerExpression().Length; i++)
            {
                var op = context.GetChild(2 * i - 1).GetText();
                var right = Visit(context.powerExpression(i));

                left = new Binary(op[0], left, right);
            }

            return left;
        }

        public override Expr VisitPowerExpression(MyGrammarParser.PowerExpressionContext context)
        {
            if (context.unaryExpression().Length == 1)
            {
                return Visit(context.unaryExpression(0));
            }

            Expr left = Visit(context.unaryExpression(0));

            for (int i = 1; i < context.unaryExpression().Length; i++)
            {
                var right = Visit(context.unaryExpression(i));
                left = new Binary('^', left, right);
            }

            return left;
        }

        public override Expr VisitUnaryExpression(MyGrammarParser.UnaryExpressionContext context)
        {
            Expr result = Visit(context.primaryExpression());

            for (int i = 0; i < context.ChildCount - 1; i++)
            {
                var op = context.GetChild(i).GetText();
                if (op == "-")
                {
                    result = new Binary('*', new Literal(-1), result);
                }
            }

            return result;
        }

        public override Expr VisitPrimaryExpression(MyGrammarParser.PrimaryExpressionContext context)
        {
            if (context.NUMBER() != null)
            {
                var numberText = context.NUMBER().GetText();
                if (double.TryParse(numberText, out var value))
                    return new Literal(value);
                return new Literal(0);
            }

            if (context.CELL_REF() != null)
            {
                var cellRef = context.CELL_REF().GetText();
                if (CellRefUtil.TryParse(cellRef, out var cr))
                {
                    return new CellRef(cr.Row, cr.Col);
                }
                throw new Exception($"Invalid cell reference: {cellRef}");
            }

            if (context.MIN() != null)
            {
                var arg1 = Visit(context.expression(0));
                var arg2 = Visit(context.expression(1));
                return new FuncCall("min", new[] { arg1, arg2 });
            }

            if (context.MAX() != null)
            {
                var arg1 = Visit(context.expression(0));
                var arg2 = Visit(context.expression(1));
                return new FuncCall("max", new[] { arg1, arg2 });
            }

            if (context.INC() != null)
            {
                var arg = Visit(context.expression(0));
                return new FuncCall("inc", new[] { arg });
            }

            if (context.DEC() != null)
            {
                var arg = Visit(context.expression(0));
                return new FuncCall("dec", new[] { arg });
            }

            if (context.expression() != null && context.expression().Length == 1)
            {
                return Visit(context.expression(0));
            }

            throw new Exception($"Unsupported primary expression: {context.GetText()}");
        }
    }
}