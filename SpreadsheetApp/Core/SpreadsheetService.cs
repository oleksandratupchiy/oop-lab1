using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SpreadsheetApp11.Expressions;
using SpreadsheetApp11.Persistence;
using SpreadsheetApp11.Parser;

namespace SpreadsheetApp11.Core
{
    public class SpreadsheetService
    {
        private Sheet _sheet = new Sheet(50, 50);
        private readonly Evaluator _eval;

        private const int MaxRows = 100_000;
        private const int MaxCols = 100_000;

        public SpreadsheetService() { _eval = new Evaluator(() => _sheet); }

        public int RowCount => _sheet.Rows;
        public int ColCount => _sheet.Cols;

        public void CreateEmpty(int rows, int cols)
        {
            _sheet = new Sheet(
                Math.Max(10, Math.Min(rows, MaxRows)),
                Math.Max(10, Math.Min(cols, MaxCols)));
        }

        public (bool Ok, string? Error) SetCell(int r, int c, string text)
        {
            var cell = _sheet.GetCell(r, c);
            cell.Error = null;
            cell.HasCachedValue = false; 

            if (string.IsNullOrWhiteSpace(text))
            {
                cell.Clear();
                return (true, null);
            }

            if (!text.StartsWith("="))
            {
                if (double.TryParse(text, out var num))
                    cell.Set(text, new Literal(num));
                else
                    cell.Set(text, new TextLiteral(text));
                return (true, null);
            }

            try
            {
                var parsedExpr = FormulaParser.Parse(text.Substring(1), _sheet);
                cell.Set(text, parsedExpr);

                return (true, null);
            }
            catch (Exception ex)
            {
                cell.Set(text, null);
                cell.Error = ex.Message;
                return (false, ex.Message);
            }
        }

        public string GetCellRaw(int r, int c)
        {
            var cell = _sheet.GetCell(r, c);
            Console.WriteLine($"🔧 GET_CELL_RAW: ({r},{c}) - RawText='{cell.RawText}', Parsed type={cell.Parsed?.GetType().Name}");
            return cell.RawText ?? "";
        }

        public string GetCellDisplay(int r, int c, DisplayMode mode)
        {
            var cell = _sheet.GetCell(r, c);
            if (cell == null) return "";

            if (mode == DisplayMode.Formulas)
                return cell.RawText ?? "";

            if (!string.IsNullOrEmpty(cell.Error))
                return "#ERR";

            if (string.IsNullOrEmpty(cell.RawText))
                return "";

            if (cell.Parsed is TextLiteral)
                return cell.RawText;

            if (cell.Parsed != null && !(cell.Parsed is Literal))
            {
                try
                {
                    var result = _eval.EvalCell(r, c);
                    return result.ToString();
                }
                catch (EvalException ex) when (ex.Message.Contains("Cycle"))
                {
                    cell.Error = "Cycle detected";
                    return "#ERR";
                }
                catch (EvalException)
                {
                    return "#ERR";
                }
                catch
                {
                    return "#ERR";
                }
            }

            if (cell.Parsed is Literal literal)
                return literal.Value.ToString();

            return "";
        }

        public void InsertRows(int at, int count)
        {
            var allCells = _sheet.GetFilled().ToList();
            int newRows = Math.Min(_sheet.Rows + count, MaxRows);
            _sheet.Resize(newRows, _sheet.Cols);

            foreach (var cell in allCells)
            {
                _sheet.ClearCell(cell.pos.r, cell.pos.c);
            }

            foreach (var cell in allCells)
            {
                int newRow = cell.pos.r;
                if (newRow >= at)
                {
                    newRow += count;
                }

                if (newRow < newRows)
                {
                    var newCell = _sheet.GetCell(newRow, cell.pos.c);

                    if (cell.cell.Parsed != null && cell.cell.RawText?.StartsWith("=") == true)
                    {
                        var shiftedExpr = ReferenceRewriter.ShiftRows(cell.cell.Parsed, at, count);

                        string newFormula = "=" + ExprFormatter.Format(shiftedExpr);
                        Console.WriteLine($"🔧 ОНОВЛЕННЯ ФОРМУЛИ: {cell.cell.RawText} -> {newFormula}");

                        newCell.Set(newFormula, shiftedExpr);
                    }
                    else
                    {
                        newCell.Set(cell.cell.RawText, cell.cell.Parsed);
                    }
                }
            }
        }


        public void DeleteRows(int at, int count)
        {
            var allCells = _sheet.GetFilled().ToList();
            int newRows = Math.Max(1, _sheet.Rows - count);

            foreach (var cell in allCells)
            {
                _sheet.ClearCell(cell.pos.r, cell.pos.c);
            }

            foreach (var cell in allCells)
            {
                int newRow = cell.pos.r;

                if (newRow >= at + count)
                {
                    newRow -= count;
                }
                else if (newRow >= at)
                {
                    continue;
                }

                if (newRow >= 0 && newRow < newRows)
                {
                    var newCell = _sheet.GetCell(newRow, cell.pos.c);

                    if (cell.cell.Parsed != null && cell.cell.RawText?.StartsWith("=") == true)
                    {
                        var shiftedExpr = ReferenceRewriter.ShiftRows(cell.cell.Parsed, at, -count);

                        string newFormula = "=" + ExprFormatter.Format(shiftedExpr);
                        Console.WriteLine($"🔧 ОНОВЛЕННЯ ФОРМУЛИ (видалення рядків): {cell.cell.RawText} -> {newFormula}");

                        newCell.Set(newFormula, shiftedExpr);
                    }
                    else
                    {
                        newCell.Set(cell.cell.RawText, cell.cell.Parsed);
                    }
                }
            }

            _sheet.Resize(newRows, _sheet.Cols);
        }

        public void InsertColumns(int at, int count)
        {
            var allCells = _sheet.GetFilled().ToList();
            int newCols = Math.Min(_sheet.Cols + count, MaxCols);
            _sheet.Resize(_sheet.Rows, newCols);

            foreach (var cell in allCells)
            {
                _sheet.ClearCell(cell.pos.r, cell.pos.c);
            }

            foreach (var cell in allCells)
            {
                int newCol = cell.pos.c;
                if (newCol >= at)
                {
                    newCol += count;
                }

                if (newCol < newCols)
                {
                    var newCell = _sheet.GetCell(cell.pos.r, newCol);

                    if (cell.cell.Parsed != null && cell.cell.RawText?.StartsWith("=") == true)
                    {
                        var shiftedExpr = ReferenceRewriter.ShiftCols(cell.cell.Parsed, at, count);

                        string newFormula = "=" + ExprFormatter.Format(shiftedExpr);
                        Console.WriteLine($"🔧 ОНОВЛЕННЯ ФОРМУЛИ (стовпці): {cell.cell.RawText} -> {newFormula}");

                        newCell.Set(newFormula, shiftedExpr);
                    }
                    else
                    {
                        newCell.Set(cell.cell.RawText, cell.cell.Parsed);
                    }
                }
            }
        }

        public void DeleteColumns(int at, int count)
        {
            var allCells = _sheet.GetFilled().ToList();
            int newCols = Math.Max(1, _sheet.Cols - count);

            foreach (var cell in allCells)
            {
                _sheet.ClearCell(cell.pos.r, cell.pos.c);
            }

            foreach (var cell in allCells)
            {
                int newCol = cell.pos.c;

                if (newCol >= at + count)
                {
                    newCol -= count;
                }
                else if (newCol >= at)
                {
                    continue;
                }

                if (newCol >= 0 && newCol < newCols)
                {
                    var newCell = _sheet.GetCell(cell.pos.r, newCol);

                    if (cell.cell.Parsed != null && cell.cell.RawText?.StartsWith("=") == true)
                    {
                        var shiftedExpr = ReferenceRewriter.ShiftCols(cell.cell.Parsed, at, -count);

                        string newFormula = "=" + ExprFormatter.Format(shiftedExpr);
                        Console.WriteLine($"🔧 ОНОВЛЕННЯ ФОРМУЛИ (видалення стовпців): {cell.cell.RawText} -> {newFormula}");

                        newCell.Set(newFormula, shiftedExpr);
                    }
                    else
                    {
                        newCell.Set(cell.cell.RawText, cell.cell.Parsed);
                    }
                }
            }

            _sheet.Resize(_sheet.Rows, newCols);
        }

        // Local JSON
        public void SaveLocal(string path)
        {
            IStorage storage = new LocalJsonStorage();
            IWorkbookSerializer ser = new JsonWorkbookSerializer();
            storage.Save(path, ser.Serialize(_sheet));
        }

        public void LoadLocal(string path)
        {
            IStorage storage = new LocalJsonStorage();
            IWorkbookSerializer ser = new JsonWorkbookSerializer();
            var data = storage.Load(path);
            _sheet = ser.Deserialize(data);
        }

        // Google Drive
        public void SaveToDrive(string nameOrId)
        {
            try
            {
                var auth = new GoogleAuthService();
                var drive = auth.GetDriveServiceAsync().GetAwaiter().GetResult();
                IStorage storage = new GoogleDriveStorage(drive);
                IWorkbookSerializer ser = new JsonWorkbookSerializer();
                storage.Save(nameOrId, ser.Serialize(_sheet));
            }
            catch (Exception ex)
            {
                throw new Exception($"Помилка збереження в Drive: {ex.Message}");
            }
        }

        public void LoadFromDrive(string nameOrId)
        {
            try
            {
                var auth = new GoogleAuthService();
                var drive = auth.GetDriveServiceAsync().GetAwaiter().GetResult();
                IStorage storage = new GoogleDriveStorage(drive);
                IWorkbookSerializer ser = new JsonWorkbookSerializer();
                var json = storage.Load(nameOrId);
                _sheet = ser.Deserialize(json);
            }
            catch (Exception ex)
            {
                throw new Exception($"Помилка завантаження з Drive: {ex.Message}");
            }
        }

        public Sheet GetSheet()
        {
            return _sheet;
        }

        public void SetSheet(Sheet sheet)
        {
            _sheet = sheet;
        }
    }
}