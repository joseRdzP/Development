using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
                //Create Folder Params//////////////////
                string folderName = "100 Alberto-98765";
                ////////////////////////////////////////

                //Rename Folder Params///////////////////////
                string entityName = "account";
                string oldRelativePath = "01 Alberto-404040";
                string newFolderName = "100 Alberto-12345";
                /////////////////////////////////////////////

                string sharePointAccessToken = SharePointHelper.GetAccessToken().Result;
                Console.WriteLine("------------------------------------------------------------------------------------------");
                Console.WriteLine("Access Token : {0}", sharePointAccessToken);
                Console.WriteLine("------------------------------------------------------------------------------------------");
                if (!string.IsNullOrEmpty(sharePointAccessToken))
                {
                    //Create Folder
                    string creationResponse = SharePointHelper.ProcessSharePointTasks(sharePointAccessToken, folderName).Result;
                    Console.WriteLine("Create Folder Response : " + creationResponse);
                    Console.WriteLine("------------------------------------------------------------------------------------------");

                    //Rename Folder
                    string renameResponse = SharePointHelper.RenameFolder(sharePointAccessToken, 
                        "https://" + SharePointCredentials.TargetHost + "/sites/" + SharePointCredentials.SPSiteName + "/",
                        $"{entityName}/{oldRelativePath}",
                        newFolderName);
                    Console.WriteLine("Rename Folder Response : " + renameResponse);
                    Console.WriteLine("------------------------------------------------------------------------------------------");
                }
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
