using System.Text.Json;
using SpreadsheetApp11.Core;
using SpreadsheetApp11.Expressions;
using SpreadsheetApp11.Parser;

namespace SpreadsheetApp11.Persistence
{
    public class JsonWorkbookSerializer : IWorkbookSerializer
    {
        private record CellDto(int r, int c, string raw);
        private record SheetDto(int rows, int cols, CellDto[] cells);

        public string Serialize(Sheet s)
        {
            var list = new System.Collections.Generic.List<CellDto>();
            foreach (var f in s.GetFilled())
                list.Add(new CellDto(f.pos.r, f.pos.c, f.cell.RawText));
            return JsonSerializer.Serialize(new SheetDto(s.Rows, s.Cols, list.ToArray()));
        }

        public Sheet Deserialize(string json)
        {
            var dto = JsonSerializer.Deserialize<SheetDto>(json);
            if (dto == null) return new Sheet(1, 1);

            var sheet = new Sheet(dto.rows, dto.cols);

            if (dto.cells != null)
            {
                foreach (var c in dto.cells)
                {
                    var cell = sheet.GetCell(c.r, c.c);
                    var text = c.raw;

                    if (string.IsNullOrEmpty(text))
                    {
                        cell.Set("", new TextLiteral(""));
                        continue;
                    }

                    try
                    {
                        Expr ast;
                        if (text.StartsWith("="))
                        {
                            var tempService = new SpreadsheetService();
                            tempService.SetSheet(sheet);
                            ast = FormulaParser.Parse(text.Substring(1), sheet);
                        }
                        else if (double.TryParse(text, out double num))
                        {
                            ast = new Literal(num);
                        }
                        else
                        {
                            ast = new TextLiteral(text);
                        }
                        cell.Set(text, ast);
                    }
                    catch (Exception)
                    {
                        cell.Set(text, new TextLiteral(text));
                    }
                }
            }
            return sheet;
        }
    }
}