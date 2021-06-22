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
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace CRMConsoleApp.Helpers
{
    public class SharePointHelper
    {
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

        public static async Task<string> ProcessSharePointTasks(string sharePointToken, string folderName)
        {
            string creationResponse = string.Empty;
            string url = "https://" + SharePointCredentials.TargetHost + "/sites/" + SharePointCredentials.SPSiteName + "/";
            using (HttpClient client = new HttpClient())
            {
                client.BaseAddress = new Uri(url);
                client.DefaultRequestHeaders.Add("Accept", "application/json; odata=verbose");
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + sharePointToken);
                string digest = GetFormDigest(sharePointToken);
                Console.WriteLine("Digest : " + digest);
                Console.WriteLine("------------------------------------------------------------------------------------------");
                try
                {
                    creationResponse = await CreateFolder(client, digest, folderName);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.WriteLine(ex.StackTrace);
                }
            }
            return creationResponse;
        }

        public static async Task<string> CreateFolder(HttpClient client, string digest, string folderName)
        {
            string creationResponse = string.Empty;

            client.DefaultRequestHeaders.Add("X-RequestDigest", digest);
            var request = CreateRequest(folderName);
            string json = JsonConvert.SerializeObject(request);
            StringContent strContent = new StringContent(json);
            strContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
            HttpResponseMessage response = await client.PostAsync("_api/web/getfolderbyserverrelativeurl('/sites/d365dev/Account')/folders", strContent);
            if (response.IsSuccessStatusCode)
            {
                creationResponse = await response.Content.ReadAsStringAsync();
            }
            else
            {
                Console.WriteLine(response.StatusCode);
                Console.WriteLine(response.ReasonPhrase);
                creationResponse = await response.Content.ReadAsStringAsync();
            }

            return creationResponse;
        }

        public static string RenameFolder(string sharePointToken, string siteUrl, string relativePath, string newFolderName)
        {
            string response = string.Empty;

            string type = GetFolderType(sharePointToken, siteUrl, relativePath);

            if (!string.IsNullOrEmpty(type))
            {
                var odataQuery = $"_api/web/GetFolderByServerRelativeUrl('{relativePath}')/ListItemAllFields";

                var url = new Uri(string.Format("{0}{1}", siteUrl, odataQuery));

                var webRequest = (HttpWebRequest)WebRequest.Create(url);
                webRequest.Headers.Add("Authorization", "Bearer " + sharePointToken);
                webRequest.Headers.Add("X-RequestDigest", GetFormDigest(sharePointToken));
                webRequest.Headers.Add("X-HTTP-Method", "MERGE");
                webRequest.Headers.Add("If-Match", "*");
                webRequest.ContentType = "application/json;odata=verbose";
                webRequest.Method = "POST";
                
                var contentToPost = @"{ '__metadata': { 'type': '" + type + "' }," +
                    " 'Title': '" + newFolderName + "'," +
                    "'FileLeafRef':'" + newFolderName + "'}";

                byte[] content = Encoding.UTF8.GetBytes(contentToPost);

                webRequest.ContentLength = content.Length;

                Stream newStream = webRequest.GetRequestStream();
                newStream.Write(content, 0, content.Length);
                newStream.Close();

                HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();

                response = webResponse.StatusDescription;
            }

            return response;
        }

        public static string GetFormDigest(string sharePointToken)
        {
            string formDigest = null;

            string resourceUrl = "https://" + SharePointCredentials.TargetHost + "/sites/" + SharePointCredentials.SPSiteName + "/_api/contextinfo";
            HttpWebRequest wreq = WebRequest.Create(resourceUrl) as HttpWebRequest;
            wreq.Headers.Add("Authorization", "Bearer " + sharePointToken);
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

        public static object CreateRequest(string folderPath)
        {
            var type = new { type = "SP.Folder" };
            var request = new { __metadata = type, ServerRelativeUrl = folderPath };
            return request;
        }

        public static string GetFolderType(string sharePointToken, string siteUrl, string relativePath)
        {
            var type = string.Empty;

            var odataQuery = $"_api/web/GetFolderByServerRelativeUrl('{relativePath}')/ListItemAllFields";

            var url = new Uri(string.Format("{0}{1}", siteUrl, odataQuery));

            var webRequest = (HttpWebRequest)WebRequest.Create(url);
            webRequest.Headers.Add("X-RequestDigest", GetFormDigest(sharePointToken));
            webRequest.Accept = "application/json;odata=verbose";
            webRequest.Headers.Add("Authorization", "Bearer " + sharePointToken);
            webRequest.Method = "GET";
            webRequest.Accept = "application/json;odata=verbose";
            webRequest.ContentLength = 0;
            webRequest.ContentType = "application/json";

            string result;
            WebResponse wresp = webRequest.GetResponse();

            using (StreamReader sr = new StreamReader(wresp.GetResponseStream()))
            {
                result = sr.ReadToEnd();
            }

            if (!string.IsNullOrEmpty(result))
            {
                var nResponse = ConvertResponse(result);

                if (nResponse != null)
                {
                    var metadata = GetResponseValue(nResponse, "__metadata");

                    if (!string.IsNullOrEmpty(metadata))
                    {
                        type = GetResponseValue(ConvertResponse(metadata, ""), "type");
                    }
                }
            }
            Console.WriteLine("Folder Type : {0}", type);
            return type;
        }

        public static Dictionary<string, string> ConvertResponse(string response, string param = "d")
        {
            var itemlist = new Dictionary<string, string>();
            dynamic e;

            if (!string.IsNullOrEmpty(param))
            {
                e = JsonConvert.DeserializeObject<dynamic>(response)[param];
            }
            else
            {
                e = JsonConvert.DeserializeObject<dynamic>(response);
            }

            if (e != null)
            {
                foreach (var item in e)
                {
                    if (item != null)
                    {
                        if (item.GetType().Name == "JObject")
                        {
                            var id = string.Empty;
                            var name = string.Empty;

                            foreach (var det in item)
                            {
                                if (det != null)
                                {
                                    var key = ((JProperty)det).Name.ToString();
                                    var val = ((JProperty)det).Value.ToString();

                                    if (key == "ServerRelativeUrl")
                                    {
                                        name = val;
                                    }

                                    if (key == "Name")
                                    {
                                        id = val;
                                    }

                                    if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(name))
                                    {
                                        itemlist.Add(id, name);
                                        break;
                                    }
                                }
                            }
                        }
                        else
                        {
                            var key = ((JProperty)item).Name.ToString();
                            var val = ((JProperty)item).Value.ToString();
                            itemlist.Add(key, val);
                        }
                    }
                }
            }
            return itemlist;
        }

        public static string GetResponseValue(Dictionary<string, string> responseList, string key)
        {
            if (responseList.ContainsKey(key))
            {
                return responseList[key];
            }
            return string.Empty;
        }
    }
}
