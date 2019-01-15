using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Google.Cloud.Vision.V1;
using BCardReader.Modals;
using System.Net.Http;
using System.Web;
using System.Text;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.Collections.Generic;

namespace BCardReader.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class GoogleReaderController : ControllerBase
    {
        private const string cognitiveTextApiKey = "aa16c83b7e274ac3a7f86028bf620b6b";


        //public const string imagepath = "E:\\dotnetimg\\";
        //public const string pythonpath = "C:\\Users\\Ujash\\AppData\\Local\\Programs\\Python\\Python37-32\\";
        // For Live
        public const string imagepath = "D:\\home\\images\\google\\";
        public const string pythonpath = "D:\\home\\python364x86\\";



        // POST api/googlereader
        [HttpPost]
        public List<UserInfo> Post([FromBody] ContactInfo value)
        {
       
            Byte[] bytes = Convert.FromBase64String(value.Image);
           
            String s2 = imagepath+"card" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".jpeg";

            System.IO.File.WriteAllBytes(s2, bytes);

            var client = ImageAnnotatorClient.Create();

            String finalOutCome = "";
            String finalresult = "";

            ProcessStartInfo start = new ProcessStartInfo();

            start.FileName = pythonpath+ "python.exe";
            start.Arguments = string.Format("{0} {1}", pythonpath+"imagesegmentation.py", "--image "+ s2 + " --paths "+ imagepath);
            start.UseShellExecute = false;
            start.CreateNoWindow = true;
            start.RedirectStandardOutput = true;
            start.RedirectStandardError = true;

            String err;
            String op;

            using (Process process = Process.Start(start))
            {

                while (!process.WaitForExit(1000)) ;
                err = process.StandardError.ReadToEnd();
                op = process.StandardOutput.ReadToEnd();

            }

            op = op.Replace("\n", string.Empty).Replace("\r", string.Empty);

            String[] imagePathArray = op.Split("|");
          
            if (System.IO.File.Exists(s2))
            {
                System.IO.File.Delete(s2);
            }

            List<UserInfo> contactDetailLst = new List<UserInfo>();

            for (int i = 0; i < imagePathArray.Length; i++)
            {

                if (!System.IO.File.Exists(imagepath + imagePathArray[i]))
                {
                    return contactDetailLst;
                }
               
                var image = Image.FromFile(imagepath + imagePathArray[i]);
                              
                String S1 = "";
                var httpresponse = client.DetectText(image);
                 
                foreach (var annotation in httpresponse)
                {
                    if (annotation.Description != null && !S1.Contains(annotation.Description))
                        if (annotation.Description.Contains("+"))
                        {
                            annotation.Description = annotation.Description.Replace(" ", "");
                        }
                            S1 = S1 + (annotation.Description) + "\n";
                }

                String[] descArray = S1.Split("\n");
                String finalS = "";
                for (int j = 0; j < descArray.Length; j++)
                {

                        if (descArray[j].StartsWith("+"))
                        {
                            descArray[j] = descArray[j].Replace(" ", "");
                        }
                        finalS = finalS + descArray[j] + "\n";
                }

                finalOutCome = finalOutCome + "{\"id\": \"" + i+1 + "\",\"language\": \"en\",\"text\":\"" + finalS + "\"},";

                if (System.IO.File.Exists(imagepath + imagePathArray[i]))
                {
                    System.IO.File.Delete(imagepath + imagePathArray[i]);
                }
            }

            finalresult = "{\"documents\":[" + finalOutCome.Substring(0, finalOutCome.Length - 1) + "]}";

                                                  
            // Process Data Using LUIS

            Task<String> t2 = GetTextAsync(finalresult);
           
            JToken token = JObject.Parse(t2.Result.ToString());

            JArray Documents = (JArray)token.SelectToken("documents");
 
            try
            {
          
                    foreach (JToken m in Documents)
                    {

                        JArray Entities = (JArray)m["entities"];

                        UserInfo contactDetail = new UserInfo();
                        contactDetail.Name = "";
                        contactDetail.Mobile = "";
                        contactDetail.Tel = "";

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

                        contactDetail.Mobile = contactDetail.Mobile.Replace("M","").Replace("\n","");

                        if (contactDetail.Mobile!=null && !contactDetail.Mobile.Equals("") && contactDetail.Mobile.Split(",").Length > 1)
                        {
                        
                            for(int j=1;j< contactDetail.Mobile.Split(",").Length;j++)
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

        private static async Task<String> GetTextAsync(String s2)
        {
            var httpclient = new HttpClient();

            var queryString = HttpUtility.ParseQueryString(string.Empty);

            // Request headers
            httpclient.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", cognitiveTextApiKey);

            var uri = "https://centralus.api.cognitive.microsoft.com/text/analytics/v2.1-preview/entities?" + queryString;

            HttpResponseMessage httpresponse;

            // Request body
            byte[] byteData = Encoding.UTF8.GetBytes(s2);

            using (var content = new ByteArrayContent(byteData))
            {
                content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                httpresponse = await httpclient.PostAsync(uri, content);
            }

            string contentString = await httpresponse.Content.ReadAsStringAsync();

            // Display the JSON response.

            return contentString;
        }
     }
}
