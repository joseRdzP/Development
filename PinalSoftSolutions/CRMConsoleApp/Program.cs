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
                string accessToken = SharePointHelper.GetAccessToken().Result;
                Console.WriteLine("SharePoint Access Token : {0}", accessToken);
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
