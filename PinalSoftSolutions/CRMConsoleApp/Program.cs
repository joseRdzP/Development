using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CRMConsoleApp.Helpers;

namespace CRMConsoleApp
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                string folderName = "01 Alberto-77777";
                string sharePointAccessToken = SharePointHelper.GetAccessToken().Result;
                Console.WriteLine("------------------------------------------------------------------------------------------");
                Console.WriteLine("SharePoint Access Token : {0}", sharePointAccessToken);
                Console.WriteLine("------------------------------------------------------------------------------------------");
                if (!string.IsNullOrEmpty(sharePointAccessToken))
                {
                    string creationResponse = SharePointHelper.ProcessSharePointTasks(sharePointAccessToken, folderName).Result;
                    Console.WriteLine("Create Folder Response : " + creationResponse);
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
