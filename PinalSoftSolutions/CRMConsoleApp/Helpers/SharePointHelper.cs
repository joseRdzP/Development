using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using CRMConsoleApp.Common;
using System.Runtime.Serialization.Json;
using System.IO;
using CRMConsoleApp.Models;
using System.Net;
using System.Web.Script.Serialization;

namespace CRMConsoleApp.Helpers
{
    public class SharePointHelper
    {
        const string msoStsUrl = "https://login.microsoftonline.com/extSTS.srf";
        const string msoHrdUrl = "https://login.microsoftonline.com/GetUserRealm.srf";
        const string spowssigninUri = "_forms/default.aspx?wa=wsignin1.0";
        const string contextInfoQuery = "_api/contextinfo";

        public static async Task<string> GetAccessToken()
        {
            SharePointResponse sharePointResponse = new SharePointResponse();
            string accessToken = string.Empty;

            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            Dictionary<string, string> keys = new Dictionary<string, string>
            {
                { "grant_type", "client_credentials" },
                { "client_id", SharePointCredentials.ClientId + "@" + SharePointCredentials.Realm },
                { "client_secret", SharePointCredentials.ClientSecret },
                { "resource", SharePointCredentials.Principal + "/" + SharePointCredentials.TargetHost + "@" + SharePointCredentials.Realm }
            };

            string httpResponse = await httpClient.PostAsync(SharePointCredentials.Uri, new FormUrlEncodedContent(keys)).Result.Content.ReadAsStringAsync();

            if (!string.IsNullOrEmpty(httpResponse))
            {
                using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(httpResponse)))
                {
                    DataContractJsonSerializer deserializer = new DataContractJsonSerializer(typeof(SharePointResponse));
                    sharePointResponse = (SharePointResponse)deserializer.ReadObject(ms);

                    if (sharePointResponse != null)
                        accessToken = sharePointResponse.access_token;
                }
            }

            return accessToken;
        }

        public static void CreateFolder(string siteUrl, string relativePath)
        {
            var odataQuery = "_api/web/folders";
            var contentToPost = @"{ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': '" + relativePath + "'}";
            byte[] content = Encoding.UTF8.GetBytes(contentToPost);
            var url = new Uri(string.Format("{0}/{1}", siteUrl, odataQuery));
            var webRequest = (HttpWebRequest)WebRequest.Create(url);
            webRequest.Headers.Add("X-RequestDigest", GetFormDigest());

        }

        public static string GetFormDigest()
        {
            string formDigest = null;

            string resourceUrl = "http://basesmc15/_api/contextinfo";
            HttpWebRequest wreq = WebRequest.Create(resourceUrl) as HttpWebRequest;
            wreq.UseDefaultCredentials = true;
            wreq.Method = "POST";
            wreq.Accept = "application/json;odata=verbose";
            wreq.ContentLength = 0;
            wreq.ContentType = "application/json";
            string result;
            WebResponse wresp = wreq.GetResponse();

            using (StreamReader sr = new StreamReader(wresp.GetResponseStream()))
            {
                result = sr.ReadToEnd();
            }

            var jss = new JavaScriptSerializer();
            var val = jss.Deserialize<Dictionary<string, object>>(result);
            var d = val["d"] as Dictionary<string, object>;
            var wi = d["GetContextWebInformation"] as Dictionary<string, object>;
            formDigest = wi["FormDigestValue"].ToString();

            return formDigest;
        }
    }
}
