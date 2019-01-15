using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.CognitiveServices.Vision.ComputerVision;
using Microsoft.Azure.CognitiveServices.Vision.ComputerVision.Models;
using System.IO;
using BCardReader.Modals;
using System.Net.Http;
using System.Web;
using System.Text;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using System.Data;
using System.Diagnostics;
using System.Collections.Generic;

namespace BCardReader.Controllers
{

    [Route("api/[controller]")]
    [ApiController]
    public class CardReaderController : ControllerBase
    {
        private const string subscriptionKey = "a774e6e5c4f9485caac89d2e74804c48";

        private const string cognitiveTextApiKey = "aa16c83b7e274ac3a7f86028bf620b6b";
            
        private const TextRecognitionMode textRecognitionMode =
            TextRecognitionMode.Printed;

        private const int numberOfCharsInOperationId = 36;

       
        // For Live
         public const string imagepath = "D:\\home\\images\\azure\\";
         public const string pythonpath = "D:\\home\\python364x86\\";



        // POST api/cardreader
        [HttpPost]
        public List<UserInfo> Post([FromBody] ContactInfo value)
        {
            ComputerVisionClient computerVision = new ComputerVisionClient(
                 new ApiKeyServiceClientCredentials(subscriptionKey),
                 new System.Net.Http.DelegatingHandler[] { });


            computerVision.Endpoint = "https://centralus.api.cognitive.microsoft.com/";


            Byte[] bytes = Convert.FromBase64String(value.Image);
            String s2 = imagepath+"card" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".jpeg";

            System.IO.File.WriteAllBytes(s2, bytes);


            ProcessStartInfo start = new ProcessStartInfo();
            start.FileName = pythonpath+"python.exe";
            start.Arguments = string.Format("{0} {1}", pythonpath+"imagesegmentation.py", "--image " + s2 + " --paths "+imagepath);
            start.UseShellExecute = false;
            start.CreateNoWindow = true;
            start.RedirectStandardOutput = true;
            start.RedirectStandardError = true; 
                                                
                                                
            DataTable dtCustomer = new DataTable("User Info");
            start.UseShellExecute = false;
            start.RedirectStandardOutput = true;
            String err;
            String op;
          
            using (Process process = Process.Start(start))
            {
           
                while (!process.WaitForExit(1000)) ;
                err = process.StandardError.ReadToEnd();
                op = process.StandardOutput.ReadToEnd();
            
            }
                                 
            op = op.Replace("\n", string.Empty).Replace("\r", string.Empty);
            
            Task<String> t2 = ExtractLocalTextAsync(computerVision, op);

            if (System.IO.File.Exists(s2))
            {
                System.IO.File.Delete(s2);
            }

        
            JToken token = JObject.Parse(t2.Result.ToString());
           
            JArray Documents = (JArray)token.SelectToken("documents");

            List<UserInfo> contactDetailLst = new List<UserInfo>();
            try
            {
           
                foreach (JToken m in Documents)
                {

                    JArray Entities = (JArray)m["entities"];

                    UserInfo contactDetail = new UserInfo();
                    contactDetail.Name = "";
                    contactDetail.Mobile = "";

                    foreach (JToken n in Entities)
                    {

                        String type = n["type"] != null && !n["type"].Equals("") ? n["type"].ToString() : null;

                        if (type != null && type.Equals("Person"))
                        {
                            contactDetail.Name = contactDetail.Name + " " + n["name"].ToString();
                        }
                        if (type != null && type.Equals("Email"))
                        {
                            contactDetail.Email = n["name"].ToString();
                        }
                        if (type != null && type.Equals("URL"))
                        {
                            contactDetail.Website = n["name"].ToString();
                        }

                        if (type != null && type.Equals("Quantity"))
                        {
                            if (n["name"].ToString().Length > 7)
                            {
                                if (contactDetail.Mobile != null && !contactDetail.Mobile.Equals(""))
                                {
                                    contactDetail.Mobile = contactDetail.Mobile + "," + n["name"].ToString();
                                }
                                else
                                {
                                    contactDetail.Mobile = contactDetail.Mobile + n["name"].ToString();
                                }
                            }
                        }
                    }

                    contactDetail.Mobile = contactDetail.Mobile.Replace("M", "").Replace("\n", "");

                    if (contactDetail.Mobile != null && !contactDetail.Mobile.Equals("") && contactDetail.Mobile.Split(",").Length > 1)
                    {

                        for (int j = 1; j < contactDetail.Mobile.Split(",").Length; j++)
                        {
                            if (contactDetail.Tel != null && !contactDetail.Tel.Equals(""))
                            {
                                contactDetail.Tel = contactDetail.Tel + "," + contactDetail.Mobile.Split(",")[j];
                            }
                            else
                            {
                                contactDetail.Tel = contactDetail.Tel + contactDetail.Mobile.Split(",")[j];
                            }
                        }
                        contactDetail.Mobile = contactDetail.Mobile.Split(",")[0];
                    }

                    contactDetailLst.Add(contactDetail);
                }
            }
            catch (Exception e)
            {

            }
            return contactDetailLst;
        }

        // Recognize text from a local image
        private static async Task<String> ExtractLocalTextAsync(
                    ComputerVisionClient computerVision, string imagePath)
        {
            String finalOutCome = "";
            String finalresult = "";

            String[] imagePathArray = imagePath.Split("|");

            for (int i = 0; i < imagePathArray.Length; i++)
            {
                if (!System.IO.File.Exists(imagepath + imagePathArray[i].Replace("\n", string.Empty).Replace("\r", string.Empty)))
                {
                    return "ERROr " + imagepath + imagePathArray[i].Replace("\n", string.Empty).Replace("\r", string.Empty);
                }
                using (Stream imageStream = System.IO.File.OpenRead(imagepath+ imagePathArray[i]))
                {
                    // Start the async process to recognize the text
                    RecognizeTextInStreamHeaders textHeaders =
                        await computerVision.RecognizeTextInStreamAsync(
                            imageStream, textRecognitionMode);

                    String result = await GetTextAsync(computerVision, textHeaders.OperationLocation);
                    finalOutCome = finalOutCome + "{\"id\": \"" + i + 1 + "\",\"language\": \"en\",\"text\":\"" + result + "\"},";

                }
                if (System.IO.File.Exists(imagepath + imagePathArray[i]))
                {
                    System.IO.File.Delete(imagepath + imagePathArray[i]);
                }
            }
            finalresult = "{\"documents\":[" + finalOutCome.Substring(0, finalOutCome.Length - 1) + "]}";
            // return finalresult;
            return await GetTextAsync(finalresult);
        }

        // Retrieve the recognized text
        private static async Task<String> GetTextAsync(
            ComputerVisionClient computerVision, string operationLocation)
        {
            // Retrieve the URI where the recognized text will be
            // stored from the Operation-Location header
            string operationId = operationLocation.Substring(
                operationLocation.Length - numberOfCharsInOperationId);

            TextOperationResult result =
                 await computerVision.GetTextOperationResultAsync(operationId);

            // Wait for the operation to complete
            int i = 0;
            int maxRetries = 10;
            String S1 = "";
            try
            {

                if (result != null && result.Status != null)
                {
                    while ((result.Status == TextOperationStatusCodes.Running ||
                            result.Status == TextOperationStatusCodes.NotStarted) && i++ < maxRetries)
                    {
                        result = await computerVision.GetTextOperationResultAsync(operationId);
                    }

                    var lines = result.RecognitionResult.Lines;

                    foreach (Line line in lines)
                    {
                        if (line.Text.Contains("+"))
                        {
                            line.Text = line.Text.Replace(" ", "");
                        }
                        S1 = S1 + line.Text + "\n";
                    }
                }
            }
            catch (Exception e)
            {
                
            }

            return S1;
        }
        private static async Task<String> GetTextAsync(String s2)
        {
            // Process Data Using LUIS

            var client = new HttpClient();

            var queryString = HttpUtility.ParseQueryString(string.Empty);

            // Request headers
            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", cognitiveTextApiKey);

            var uri = "https://centralus.api.cognitive.microsoft.com/text/analytics/v2.1-preview/entities?" + queryString;

            HttpResponseMessage response;

            // Request body
            byte[] byteData = Encoding.UTF8.GetBytes(s2);

            using (var content = new ByteArrayContent(byteData))
            {
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                response = await client.PostAsync(uri, content);
            }

            string contentString = await response.Content.ReadAsStringAsync();

            // Display the JSON response.

            return contentString;
        }
    }
}
