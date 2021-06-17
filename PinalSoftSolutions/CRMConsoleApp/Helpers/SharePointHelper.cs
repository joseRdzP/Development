using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using CRMConsoleApp.Common;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.IO;
using Newtonsoft.Json;

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
                { "client_id", "b62347f1-aee8-44f2-acde-031a6bada525@6331a538-74db-4756-8dce-572f8c8ceb84" },
                { "client_secret", "9+/YmIYDmunmZGp+sU1fw6wqHKUDanphfkTKVzgvnw4=" },
                { "resource", "00000003-0000-0ff1-ce00-000000000000/dynamics365cedevs.sharepoint.com@6331a538-74db-4756-8dce-572f8c8ceb84" }
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

        [DataContract]
        public class SharePointResponse
        {
            [DataMember]
            public string token_type { get; set; }
            [DataMember]
            public string expires_in { get; set; }
            [DataMember]
            public string not_before { get; set; }
            [DataMember]
            public string expires_on { get; set; }
            [DataMember]
            public string resource { get; set; }
            [DataMember]
            public string access_token { get; set; }
        }
    }
}
