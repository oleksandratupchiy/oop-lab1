using System.IO;


namespace SpreadsheetApp11.Persistence
{
    public class LocalJsonStorage : IStorage
    {
        public void Save(string path, string content) => File.WriteAllText(path, content);
        public string Load(string path) => File.ReadAllText(path);
    }
}