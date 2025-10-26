using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadsheetApp11.Persistence
{
    public class GoogleAuthService : IGoogleAuthService
    {
        private readonly string[] _scopes = { DriveService.Scope.Drive };
        private const string ApplicationName = "Spreadsheet App";

        private const string CredentialsPath = "client_secrets.json";

        private const string TokenCacheFolder = "DriveTokenCache";

        private DriveService? _driveService;

        public async Task<DriveService> GetDriveServiceAsync()
        {
            if (_driveService != null)
            {
                return _driveService;
            }

            UserCredential credential;

            using (var stream = new FileStream(CredentialsPath, FileMode.Open, FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    _scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(TokenCacheFolder, true));
            }

            _driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return _driveService;
        }

        public async Task SignOutAsync()
        {
            if (_driveService != null)
            {
                _driveService.Dispose();
                _driveService = null;
            }

            var dataStore = new FileDataStore(TokenCacheFolder, true);
            await dataStore.ClearAsync();

            if (Directory.Exists(TokenCacheFolder))
            {
                Directory.Delete(TokenCacheFolder, true);
            }
        }
    }
}