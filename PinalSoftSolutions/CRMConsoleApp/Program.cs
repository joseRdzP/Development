using System;
using System.Threading.Tasks;
using zCRMConsoleApp.Common;
using zCRMConsoleApp.Helpers;

namespace zCRMConsoleApp
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                //Call OData Legacy API
                //PluginHelper.GetAccountData();
                PluginHelper.GetAccountDataAzure();

                /*
                //Get SP Configurations
                string spClientId = SharePointCredentials.ClientId;
                string spClientSecret = SharePointCredentials.ClientSecret;
                string spRealm = SharePointCredentials.Realm;
                string spPrincipal = SharePointCredentials.Principal;
                string spTargetHost = SharePointCredentials.TargetHost;
                string spSiteUrl = $"https://{spTargetHost}/sites/{SharePointCredentials.SPSiteName}";
                string spUri = SharePointCredentials.Uri;

                //Create Folder Params//////////////////
                string folderName = "Test 001 Alberto-123456";
                string extraFolders = "Order Agreements & Amendments,Quotes & Proposals,Termination Agreements";
                ////////////////////////////////////////

                //Rename Folder Params///////////////////////
                string entityName = "account";
                string oldRelativePath = "01 Alberto-567890";
                string newFolderName = "03 Alberto-333333";
                /////////////////////////////////////////////

                Console.WriteLine("------------------------------------------------------------------------------------------");

                //Getting the SP Access Token
                string spAccessToken = SharePointHelper.GetSpAccessToken(spClientId, spRealm, spClientSecret, spPrincipal, spTargetHost, spUri).Result;

                if (!string.IsNullOrEmpty(spAccessToken))
                {
                    //Getting the SharePoint Service instance
                    SharePointService spService = new SharePointService(spAccessToken, spSiteUrl);

                    if (spService != null)
                    {
                        //Create Folder
                        Console.WriteLine("Creating the Main Folder : " + $"{entityName}/{folderName}");
                        string creationResponse = spService.CreateFolder($"{entityName}/{folderName}");
                        Console.WriteLine("Create Folder Response : " + creationResponse);
                        Console.WriteLine("------------------------------------------------------------------------------------------");

                        //Create Extra Folders
                        Console.WriteLine("Creating Sub-Folders...");
                        var folderList = extraFolders.Split(',');
                        for (int i = 0; i < folderList.Length; i++)
                        {
                            string currentExtraFolderName = folderList[i];
                            Console.WriteLine("- Sub-Folder #" + i + " - " + $"{entityName}/{folderName}/{currentExtraFolderName}");
                            string extraFolderCreationResponse = spService.CreateFolder($"{entityName}/{folderName}/{currentExtraFolderName}");
                            Console.WriteLine("Create Extra Folder Response #" + i + " - " + extraFolderCreationResponse);
                            Console.WriteLine("------------------------------------------------------------------------------------------");
                        }
                        
                        //Rename Folder
                        string renameResponse = spService.RenameFolder($"{entityName}/{oldRelativePath}", newFolderName);
                        Console.WriteLine("Rename Folder Response : " + renameResponse);
                        Console.WriteLine("------------------------------------------------------------------------------------------");
                    }
                }
                */
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
