using System;
using System.IO;
using System.Net;

namespace zCRMConsoleApp.Helpers
{
    public static class PluginHelper
    {
        public static void GetAccountData()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

            var odataQuery = "https://accountmgt.opentechalliance.com/odata/Account?$top=1";

            var webRequest = (HttpWebRequest)WebRequest.Create(odataQuery);
            webRequest.Headers.Add("AccessToken", "095c1d5b-545c-4891-8ab2-c02057653689");
            webRequest.ContentType = "application/json; charset=utf-8";
            webRequest.Method = "GET";

            HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

            if (webResponse != null)
            {
                var responseStream = webResponse.GetResponseStream();
                if (responseStream != null)
                {
                    var reader = new StreamReader(responseStream);
                    string receiveContent = reader.ReadToEnd();
                    reader.Close();

                    Console.WriteLine(receiveContent);
                }
            }
        }
    }
}
