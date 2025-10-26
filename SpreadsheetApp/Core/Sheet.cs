using System.Collections.Generic;

namespace SpreadsheetApp11.Core
{
    public class Sheet
    {
        private readonly Dictionary<(int r, int c), Cell> _cells = new();
        public int Rows { get; private set; }
        public int Cols { get; private set; }

        public Sheet(int rows, int cols)
        {
            Rows = rows;
            Cols = cols;
        }

        public Cell GetCell(int r, int c)
        {
            if (r < 0 || r >= Rows || c < 0 || c >= Cols)
                return new Cell();

            if (!_cells.TryGetValue((r, c), out var cell))
            {
                cell = new Cell();
                _cells[(r, c)] = cell;
            }

            return cell;
        }

        public IEnumerable<((int r, int c) pos, Cell cell)> GetFilled()
        {
            foreach (var kv in _cells)
            {
                if (kv.Value != null && !string.IsNullOrEmpty(kv.Value.RawText))
                    yield return (kv.Key, kv.Value);
            }
        }

        public void Resize(int newRows, int newCols)
        {
            Rows = newRows;
            Cols = newCols;
        }

        public void ClearCell(int r, int c)
        {
            if (r >= 0 && r < Rows && c >= 0 && c < Cols)
                _cells.Remove((r, c));
        }
    }
}