using System.Threading.Tasks;

namespace CRMConsoleApp.Interfaces
{
    public interface ISharePointService
    {
        string CreateFolder(string relativePath);

        string RenameFolder(string relativePath, string newFolderName);
    }
}
