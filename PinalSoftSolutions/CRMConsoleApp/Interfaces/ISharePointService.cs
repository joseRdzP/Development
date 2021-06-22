namespace CRMConsoleApp.Interfaces
{
    public interface ISharePointService
    {
        string CreateFolder(string relativePath);

        string RenameFolder(string relativePath, string newFolderName);

        string GetFolderType(string sharePointToken, string relativePath);
    }
}
