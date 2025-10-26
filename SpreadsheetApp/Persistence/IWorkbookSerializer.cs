using SpreadsheetApp11.Core;


namespace SpreadsheetApp11.Persistence
{
    public interface IWorkbookSerializer
    {
        string Serialize(Sheet s);
        Sheet Deserialize(string json);
    }
}