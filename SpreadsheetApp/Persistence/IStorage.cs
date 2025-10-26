namespace SpreadsheetApp11.Persistence
{
    public interface IStorage
    {
        void Save(string pathOrId, string content);
        string Load(string pathOrId);
    }
}
