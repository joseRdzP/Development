using System;
using System.Collections.Generic;
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
    public static class SharePointHelper
    {
        public static async Task<string> GetSpAccessToken(string spClientId, string spRealm, string spClientSecret, string spPrincipal, string spTargetHost)
        {
            SharePointResponse sharePointResponse = new SharePointResponse();
            string accessToken = string.Empty;

            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            Dictionary<string, string> keys = new Dictionary<string, string>
            {
                { "grant_type", "client_credentials" },
                { "client_id", spClientId + "@" + spRealm },
                { "client_secret", spClientSecret },
                { "resource", spPrincipal + "/" + spTargetHost + "@" + spRealm }
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

        public static string GetFormDigest(string spAccessToken, string spSiteUrl)
        {
            string formDigest = null;

            string resourceUrl = spSiteUrl + "/_api/contextinfo";
            HttpWebRequest wreq = WebRequest.Create(resourceUrl) as HttpWebRequest;
            wreq.Headers.Add("Authorization", "Bearer " + spAccessToken);
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

        public static string GetFolderType(string spAccessToken, string spSiteUrl, string relativePath)
        {
            var type = string.Empty;
            string result = string.Empty;
            var odataQuery = $"_api/web/GetFolderByServerRelativeUrl('{relativePath}')/ListItemAllFields";
            var url = new Uri(string.Format("{0}/{1}", spSiteUrl, odataQuery));

            var webRequest = (HttpWebRequest)WebRequest.Create(url);
            webRequest.Headers.Add("X-RequestDigest", GetFormDigest(spAccessToken, spSiteUrl));
            webRequest.Accept = "application/json;odata=verbose";
            webRequest.Headers.Add("Authorization", "Bearer " + spAccessToken);
            webRequest.Method = "GET";
            webRequest.Accept = "application/json;odata=verbose";
            webRequest.ContentLength = 0;
            webRequest.ContentType = "application/json";
  
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
