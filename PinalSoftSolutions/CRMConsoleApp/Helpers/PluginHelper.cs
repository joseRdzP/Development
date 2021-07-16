using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using zCRMConsoleApp.Models;
using System.Runtime.Serialization.Json;
using System.Text;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Web.Script.Serialization;
using System.Runtime.Serialization;
using System.Reflection;

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

        public static void GetAccountDataAzure(string apiUrl)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

            var webRequest = (HttpWebRequest)WebRequest.Create(apiUrl);
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

        public static void RetrieveMultipleResultsForView(string legacyDataUrl)
        {
			try
			{
				EntityCollection collection = new EntityCollection();
				Console.WriteLine("RetrieveMultiplePlugin : Retrieving Facility DIDs...");
				try
				{
					using (HttpClient client = new HttpClient())
					{
						client.Timeout = TimeSpan.FromMilliseconds(15000);
						client.DefaultRequestHeaders.ConnectionClose = true;
						client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
						client.DefaultRequestHeaders.Add("AccessToken", "095c1d5b-545c-4891-8ab2-c02057653689");
						var response = client.GetAsync(legacyDataUrl).Result;
						response.EnsureSuccessStatusCode();

						var responseText = response.Content.ReadAsStringAsync().Result;
						
						if (!string.IsNullOrEmpty(responseText))
                        {
							Console.WriteLine("RetrieveMultiplePlugin : Facility DIDs Retrieved!");

							Console.WriteLine("RetrieveMultiplePlugin : Deserializing Facility DIDs Response...");
							JObject facilityDIDJObject = JObject.Parse(responseText);

							if (facilityDIDJObject != null)
                            {
								JToken facilityDIDJToken = facilityDIDJObject["value"];

								if (facilityDIDJToken != null)
                                {
									List<FacilityDID> facilityDIDs = JsonConvert.DeserializeObject<List<FacilityDID>>(facilityDIDJToken.ToString());
									Console.WriteLine("Number of FacilityDIDs = {0}", facilityDIDs.Count);

									if (facilityDIDs.Count > 0)
                                    {
										foreach (FacilityDID facilityDID in facilityDIDs)
                                        {
											Entity e = new Entity("ota_legacyfacilitydid");

											Dictionary<string, string> facilityDIDDictionary = JObject.FromObject(facilityDID).ToObject<Dictionary<string, string>>();

											if (facilityDIDDictionary.Count > 0)
											{
												foreach (dynamic keyValuePair in facilityDIDDictionary)
												{
													string key = keyValuePair.Key;
													object val = keyValuePair.Value;
													Console.WriteLine("=====> Processing Column Name : {0} with Value = {1}", key, val);
													e.Attributes[key] = val;
												}
												collection.Entities.Add(e);
											}
											//Type fieldsType = typeof(FacilityDID);

											//PropertyInfo[] props = fieldsType.GetProperties(BindingFlags.Public | BindingFlags.Instance);

											//for (int i = 0; i < props.Length; i++)
											//{

											//}
										}
									}
								}					
							}						
						}									

						//Console.WriteLine("RetrieveMultiplePlugin : Sending the Entity Collection to the Output Parameters...");
						//context.OutputParameters["BusinessEntityCollection"] = collection;
						//Console.WriteLine("RetrieveMultiplePlugin : Output Parameters operation is Done.");
						//throw new InvalidPluginExecutionException("Ticket Sources Retrieved.");
					}
				}
				catch (AggregateException aex)
				{
					Console.WriteLine("Inner Exceptions: ");
					foreach (Exception ex in aex.InnerExceptions)
					{
						Console.WriteLine("Exception: {0}", ex.ToString());
					}
					throw new InvalidPluginExecutionException(string.Format(CultureInfo.InvariantCulture,
						"An exception occurred while retrieving Ticket Sources.", aex));
				}
				Console.WriteLine("Collection Items = {0}", collection.Entities.Count);
			}
			catch (Exception e)
			{
				Console.WriteLine("Exception: {0}", e.ToString());
				throw;
			}
		}
	}
}
