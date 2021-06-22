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
using CRMConsoleApp.Interfaces;

namespace CRMConsoleApp.Helpers
{
    public class SharePointHelper : ISharePointService
    {
        private string _spSiteUrl;
        private string _spClientId;
        private string _spClientSecret;
        private string _spRealm;
        private string _spPrincipal;
        private string _spTargetHost;
        private string _digest;

        public SharePointHelper(string spSiteUrl, string spClientId, string spClientSecret, string spRealm, string spPrincipal, string spTargetHost)
        {
            _spSiteUrl = spSiteUrl;
            _spClientId = spClientId;
            _spClientSecret = spClientSecret;
            _spRealm = spRealm;
            _spPrincipal = spPrincipal;
            _spTargetHost = spTargetHost;
            _digest = GetFormDigest(GetAccessToken().Result);
        }

        public async Task<string> GetAccessToken()
        {
            SharePointResponse sharePointResponse = new SharePointResponse();
            string accessToken = string.Empty;

            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            Dictionary<string, string> keys = new Dictionary<string, string>
            {
                { "grant_type", "client_credentials" },
                { "client_id", _spClientId + "@" + _spRealm },
                { "client_secret", _spClientSecret },
                { "resource", _spPrincipal + "/" + _spTargetHost + "@" + _spRealm }
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

        public async Task<string> CreateFolder(string relativePath)
        {
            string createFolderResponse = string.Empty;

            try
            {
                string accessToken = await GetAccessToken();

                if (!string.IsNullOrEmpty(accessToken))
                {
                    var odataQuery = "_api/web/folders";
                    var contentToPost = @"{ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': '" + relativePath + "'}";
                    byte[] content = Encoding.UTF8.GetBytes(contentToPost);

                    var url = new Uri(string.Format("{0}/{1}", _spSiteUrl, odataQuery));

                    var webRequest = (HttpWebRequest)WebRequest.Create(url);
                    webRequest.Headers.Add("Authorization", "Bearer " + accessToken);
                    webRequest.Headers.Add("X-RequestDigest", _digest);
                    webRequest.ContentLength = content.Length;
                    webRequest.ContentType = "application/json;odata=verbose";
                    webRequest.Method = "POST";

                    Stream newStream = webRequest.GetRequestStream();
                    newStream.Write(content, 0, content.Length);
                    newStream.Close();

                    HttpWebResponse webResponse = (HttpWebResponse)webRequest.GetResponse();
                    createFolderResponse = webResponse.StatusDescription;
                }
            }
            catch (Exception ex)
            {
                createFolderResponse = ex.Message;
                Console.WriteLine(ex.Message);
            }

            return createFolderResponse;
        }

        public async Task<string> RenameFolder(string relativePath, string newFolderName)
        {
            string renameFolderResponse = string.Empty;
            string accessToken = await GetAccessToken();

            if (!string.IsNullOrEmpty(accessToken))
            {
                string type = GetFolderType(accessToken, relativePath);

                if (!string.IsNullOrEmpty(type))
                {
                    var odataQuery = $"_api/web/GetFolderByServerRelativeUrl('{relativePath}')/ListItemAllFields";

                    var url = new Uri(string.Format("{0}/{1}", _spSiteUrl, odataQuery));

                    var webRequest = (HttpWebRequest)WebRequest.Create(url);
                    webRequest.Headers.Add("Authorization", "Bearer " + accessToken);
                    webRequest.Headers.Add("X-RequestDigest", _digest);
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

                    renameFolderResponse = webResponse.StatusDescription;
                }
            }

            return renameFolderResponse;
        }

        public string GetFormDigest(string sharePointToken)
        {
            string formDigest = null;

            string resourceUrl = _spSiteUrl + "/_api/contextinfo";
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

        public string GetFolderType(string sharePointToken, string relativePath)
        {
            var type = string.Empty;
            string result = string.Empty;
            var odataQuery = $"_api/web/GetFolderByServerRelativeUrl('{relativePath}')/ListItemAllFields";
            var url = new Uri(string.Format("{0}/{1}", _spSiteUrl, odataQuery));

            var webRequest = (HttpWebRequest)WebRequest.Create(url);
            webRequest.Headers.Add("X-RequestDigest", _digest);
            webRequest.Accept = "application/json;odata=verbose";
            webRequest.Headers.Add("Authorization", "Bearer " + sharePointToken);
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

        public Dictionary<string, string> ConvertResponse(string response, string param = "d")
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

        public string GetResponseValue(Dictionary<string, string> responseList, string key)
        {
            if (responseList.ContainsKey(key))
            {
                return responseList[key];
            }
            return string.Empty;
        }
    }
}
