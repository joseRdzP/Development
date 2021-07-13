using System;
using System.Text;
using System.IO;
using System.Net;
using zCRMConsoleApp.Interfaces;
using zCRMConsoleApp.Helpers;

namespace zCRMConsoleApp.Common
{
    public class SharePointService : ISharePointService
    {
        private string _spAccessToken;
        private string _spSiteUrl;
        private string _digest;

        public SharePointService(string spAccessToken, string spSiteUrl)
        {
            _spAccessToken = spAccessToken;
            _spSiteUrl = spSiteUrl;
            _digest = SharePointHelper.GetFormDigest(_spAccessToken, _spSiteUrl);
        }

        public string CreateFolder(string relativePath)
        {
            string createFolderResponse = string.Empty;

            try
            {
                if (!string.IsNullOrEmpty(_spAccessToken))
                {
                    var odataQuery = "_api/web/folders";
                    var contentToPost = @"{ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': '" + relativePath + "'}";
                    byte[] content = Encoding.UTF8.GetBytes(contentToPost);

                    var url = new Uri(string.Format("{0}/{1}", _spSiteUrl, odataQuery));

                    var webRequest = (HttpWebRequest)WebRequest.Create(url);
                    webRequest.Headers.Add("Authorization", "Bearer " + _spAccessToken);
                    webRequest.Headers.Add("X-RequestDigest", _digest);
                    webRequest.ContentLength = content.Length;
                    webRequest.ContentType = "application/json;odata=verbose";
                    webRequest.Method = "POST";
                    webRequest.AllowAutoRedirect = false;
                    webRequest.UserAgent = "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)";
                    webRequest.Accept = "application/json;odata=verbose";

                    Stream dataStream = webRequest.GetRequestStream();
                    dataStream.Write(content, 0, content.Length);
                    dataStream.Close();

                    WebResponse response = webRequest.GetResponse();
                    createFolderResponse = ((HttpWebResponse)response).StatusDescription;

                    using (dataStream = response.GetResponseStream())
                    {
                        StreamReader reader = new StreamReader(dataStream);
                        string responseFromServer = reader.ReadToEnd();
                    }

                    response.Close();
                }
            }
            catch (Exception ex)
            {
                createFolderResponse = ex.Message;
                Console.WriteLine(ex.Message);
            }

            return createFolderResponse;
        }

        public string RenameFolder(string relativePath, string newFolderName)
        {
            string renameFolderResponse = string.Empty;

            if (!string.IsNullOrEmpty(_spAccessToken))
            {
                string type = SharePointHelper.GetFolderType(_spAccessToken, _spSiteUrl, relativePath);

                if (!string.IsNullOrEmpty(type))
                {
                    var odataQuery = $"_api/web/GetFolderByServerRelativeUrl('{relativePath}')/ListItemAllFields";

                    var url = new Uri(string.Format("{0}/{1}", _spSiteUrl, odataQuery));

                    var webRequest = (HttpWebRequest)WebRequest.Create(url);
                    webRequest.Headers.Add("Authorization", "Bearer " + _spAccessToken);
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
    }
}
