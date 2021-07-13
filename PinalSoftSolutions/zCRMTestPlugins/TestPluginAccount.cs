using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace zCRMTestPlugins
{
    public class TestPluginAccount : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracingService.Trace("TestPluginAccount : Initiating Plugin...");

            // Obtain the execution context from the service provider.  
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.  
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.  
                Entity entity = (Entity)context.InputParameters["Target"];

                // Obtain the organization service reference which you will need for  
                // web service calls.  
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try
                {
                    // Plug-in business logic goes here. 
                    tracingService.Trace("TestPluginAccount : Executing Plugin...");

                    GetAccountData3(tracingService);
                }

                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in TestPluginAccount.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("TestPluginAccount: {0}", ex.ToString());
                    throw;
                }
            }
        }

        public static void GetAccountData1(ITracingService tracer)
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

                    tracer.Trace(receiveContent);
                }
            }
        }

        public static void GetAccountData2(ITracingService tracer)
        {
            var requestUrl = "https://accountmgt.opentechalliance.com/odata/FacilityDID";
            var accessToken = "095c1d5b-545c-4891-8ab2-c02057653689";

            var webClient = new WebClient();
            webClient.Headers[HttpRequestHeader.ContentType] = "application/json";
            webClient.Headers["AccessToken"] = accessToken;

            var response = webClient.DownloadString(requestUrl);

            if (response != null)
            {
                tracer.Trace(response);
            }
        }

        public static void GetAccountData3(ITracingService tracer)
        {
            var odataQuery = "https://accountmgt.opentechalliance.com/odata/Account?$top=1";

            HttpClient client = new HttpClient();

            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, odataQuery)
            {
                Content = new StringContent("", Encoding.UTF8)
            };
            request.Content.Headers.Remove("Content-Type");
            request.Content.Headers.TryAddWithoutValidation("Content-Type", "application/x-www-form-urlencoded");

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            var responseMessage = client.SendAsync(request);
            var jsonResponseString = responseMessage.GetAwaiter().GetResult();
        }
    }
}
