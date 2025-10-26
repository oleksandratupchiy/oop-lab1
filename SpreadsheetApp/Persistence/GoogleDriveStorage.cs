using Google.Apis.Drive.v3;
using Google.Apis.Upload;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Linq;

namespace SpreadsheetApp11.Persistence
{
    public class GoogleDriveStorage : IStorage
    {
        private readonly DriveService _driveService;
        private const string JsonMimeType = "application/json";

        public GoogleDriveStorage(DriveService driveService)
        {
            _driveService = driveService;
        }

        public void Save(string nameOrId, string content)
        {
            SaveAsync(nameOrId, content).GetAwaiter().GetResult();
        }

        public string Load(string nameOrId)
        {
            return LoadAsync(nameOrId).GetAwaiter().GetResult();
        }

        private async Task SaveAsync(string nameOrId, string content)
        {
            try
            {
                var existingFile = await FindFileByName(nameOrId);

                byte[] byteArray = System.Text.Encoding.UTF8.GetBytes(content);
                using var stream = new MemoryStream(byteArray);

                if (existingFile != null)
                {

                    var updateRequest = _driveService.Files.Update(
                        new Google.Apis.Drive.v3.Data.File(),
                        existingFile.Id,
                        stream,
                        JsonMimeType); 

                    var result = await updateRequest.UploadAsync();
                    if (result.Status != UploadStatus.Completed)
                        throw new Exception($"Google Drive Update failed: {result.Exception?.Message ?? result.Status.ToString()}");
                }
                else
                {
                    var fileMetadata = new Google.Apis.Drive.v3.Data.File()
                    {
                        Name = nameOrId,
                        MimeType = JsonMimeType 
                    };

                    var createRequest = _driveService.Files.Create(fileMetadata, stream, JsonMimeType); // Використовуємо JSON MIME тип
                    createRequest.Fields = "id";

                    var result = await createRequest.UploadAsync();
                    if (result.Status != UploadStatus.Completed)
                        throw new Exception($"Google Drive Create failed: {result.Exception?.Message ?? result.Status.ToString()}");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Помилка збереження в Drive: {ex.Message}");
            }
        }

        private async Task<string> LoadAsync(string nameOrId)
        {
            var file = await FindFileByName(nameOrId);

            if (file == null)
            {
                throw new FileNotFoundException($"Файл '{nameOrId}' не знайдено на Google Drive.");
            }

            var request = _driveService.Files.Get(file.Id);
            using var stream = new MemoryStream();
            await request.DownloadAsync(stream);

            stream.Position = 0;
            using var reader = new StreamReader(stream, System.Text.Encoding.UTF8);
            return await reader.ReadToEndAsync();
        }

        private async Task<Google.Apis.Drive.v3.Data.File> FindFileByName(string name)
        {
            try
            {
                var listRequest = _driveService.Files.List();
                listRequest.Q = $"name='{name}' and trashed=false";
                listRequest.Fields = "files(id, name)";
                listRequest.Spaces = "drive";

                var files = await listRequest.ExecuteAsync();

                return files.Files.FirstOrDefault();
            }
            catch (Exception ex)
            {
                throw new Exception($"Помилка пошуку файлу: {ex.Message}");
            }
        }

        public async Task<List<string>> ListAllFiles()
        {
            try
            {
                var listRequest = _driveService.Files.List();
                listRequest.Fields = "files(id, name, createdTime)";
                listRequest.Spaces = "drive";
                listRequest.OrderBy = "createdTime desc";
                listRequest.PageSize = 100;

                var files = await listRequest.ExecuteAsync();
                var result = new List<string>();

                foreach (var file in files.Files)
                {
                    result.Add($"{file.Name} (ID: {file.Id})");
                }

                return result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Помилка отримання списку файлів: {ex.Message}");
            }
        }
    }
}