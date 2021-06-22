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

                //Getting the SharePoint Service Instance
                SharePointHelper sharePointService = new SharePointHelper(spSiteUrl, spClientId, spClientSecret, spRealm, spPrincipal, spTargetHost);

                //Create Folder
                //string creationResponse = SharePointHelper.CreateSharePointFolder(sharePointAccessToken, folderName).Result;
                string creationResponse = sharePointService.CreateFolder($"{entityName}/{folderName}");
                Console.WriteLine("Create Folder Response : " + creationResponse);
                Console.WriteLine("------------------------------------------------------------------------------------------");

                //Create Extra Folders
                string[] folderList = extraFolders.Split(char.Parse(","));
                foreach (var extraFolderName in folderList)
                {
                    string extraFolderCreationResponse = sharePointService.CreateFolder($"{entityName}/{folderName}/{extraFolderName}");
                    Console.WriteLine("Create Extra Folder Response : " + extraFolderCreationResponse);
                    Console.WriteLine("------------------------------------------------------------------------------------------");
                }

                //Rename Folder
                string renameResponse = sharePointService.RenameFolder($"{entityName}/{oldRelativePath}", newFolderName);
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
