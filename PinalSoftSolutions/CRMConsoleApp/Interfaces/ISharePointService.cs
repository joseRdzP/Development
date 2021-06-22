using System.Threading.Tasks;

namespace CRMConsoleApp.Interfaces
{
    public interface ISharePointService
    {
        Task<string> CreateFolder(string relativePath);

        Task<string> RenameFolder(string relativePath, string newFolderName);

        string GetFolderType(string sharePointToken, string relativePath);
    }
}
