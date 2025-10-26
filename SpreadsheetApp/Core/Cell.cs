using SpreadsheetApp11.Expressions;

namespace SpreadsheetApp11.Core
{
    public class Cell
    {
        public string RawText { get; private set; } = string.Empty;
        public Expr? Parsed { get; private set; }
        public double CachedValue { get; set; } = 0;
        public bool HasCachedValue { get; set; } = false;
        public string? Error { get; set; }

        public void Set(string raw, Expr? ast)
        {
            Console.WriteLine($"🔧 CELL.SET: raw='{raw}', ast type={ast?.GetType().Name}");
            RawText = raw ?? "";
            Parsed = ast;
            CachedValue = 0;
            HasCachedValue = false;
            Error = null;
            Console.WriteLine($"🔧 CELL.SET: After - RawText='{RawText}', Parsed type={Parsed?.GetType().Name}");
        }

        public void Clear()
        {
            RawText = "";
            Parsed = null;
            CachedValue = 0;
            HasCachedValue = false;
            Error = null;
        }
    }
}