using System;
using System.Threading.Tasks;
using CRMConsoleApp.Common;
using CRMConsoleApp.Helpers;

namespace CRMConsoleApp
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                //Get SP Configurations
                string spClientId = SharePointCredentials.ClientId;
                string spClientSecret = SharePointCredentials.ClientSecret;
                string spRealm = SharePointCredentials.Realm;
                string spPrincipal = SharePointCredentials.Principal;
                string spTargetHost = SharePointCredentials.TargetHost;
                string spSiteUrl = $"https://{spTargetHost}/sites/{SharePointCredentials.SPSiteName}";

                //Create Folder Params//////////////////
                string folderName = "00 Alberto-123456";
                string extraFolders = "Order Agreements & Amendments,Quotes & Proposals,Termination Agreements";
                ////////////////////////////////////////

                //Rename Folder Params///////////////////////
                string entityName = "account";
                string oldRelativePath = "300 Alberto-123456";
                string newFolderName = "300 Alberto-111111";
                /////////////////////////////////////////////

                Console.WriteLine("------------------------------------------------------------------------------------------");

                //Getting the SharePoint Service Instance
                SharePointHelper sharePointService = new SharePointHelper(spSiteUrl, spClientId, spClientSecret, spRealm, spPrincipal, spTargetHost);

                //Create Folder
                string creationResponse = await sharePointService.CreateFolder($"{entityName}/{folderName}");
                Console.WriteLine("Create Folder Response : " + creationResponse);
                Console.WriteLine("------------------------------------------------------------------------------------------");
                await Task.Delay(5000);

                //Create Extra Folders
                var folderList = extraFolders.Split(',');
                for (int i = 0; i < folderList.Length; i++)
                {
                    string currentExtraFolderName = folderList[i];
                    string extraFolderCreationResponse = await sharePointService.CreateFolder($"{entityName}/{folderName}/{currentExtraFolderName}");
                    Console.WriteLine("Create Extra Folder Response : " + extraFolderCreationResponse);
                    Console.WriteLine("------------------------------------------------------------------------------------------");
                    await Task.Delay(2000);
                }

                //Rename Folder
                string renameResponse = await sharePointService.RenameFolder($"{entityName}/{oldRelativePath}", newFolderName);
                Console.WriteLine("Rename Folder Response : " + renameResponse);
                Console.WriteLine("------------------------------------------------------------------------------------------");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught - " + ex.Message);
            }
            finally
            {
                Console.WriteLine("Press Enter to Exit from the Console App...");
                Console.ReadKey();
            }
        }
    }
}
