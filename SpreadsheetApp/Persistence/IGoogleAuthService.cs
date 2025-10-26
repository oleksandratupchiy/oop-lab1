using System.Threading.Tasks;
using Google.Apis.Drive.v3;


namespace SpreadsheetApp11.Persistence
{
    public interface IGoogleAuthService
    {
        Task<DriveService> GetDriveServiceAsync();
        Task SignOutAsync();
    }
}