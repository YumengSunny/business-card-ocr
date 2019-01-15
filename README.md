
*****************
Deployement Steps
*****************

1) Go to azzure accounnt. Click on Create new resource. Select Cognitive service and complete the process by filling simple form. You will be provided key to access this service.
   Copy that key and replace with value of variable subscriptionKey in CardReaderController.cs.

2) Go to azzure accounnt. Click on Create new resource. Select Text Analytics and complete the process by filling simple form. You will be provided key to access this service.
   Copy that key and replace with value of variable cognitiveTextApiKey in CardReaderController.cs and GoogleReaderController.cs.  
 
3) Keep json file of google vision's api to D:\home\myapp\apikeys.json and give this path in Settings >Application settings> Application settings of deployeed app.
 
4) Make sure that python is installed with open-cv and argparse packages. Replace python's path with value of variable pythonpath in CardReaderController.cs and GoogleReaderController.cs. 

5) Create folder D:\\home\\images\\google\\ with full permission. 

6) Create folder D:\\home\\images\\azure\\ with full permission. 
   
7) Follow the steps given in link to deployee the web application in azure.
https://docs.microsoft.com/en-us/azure/app-service/app-service-web-get-started-dotnet

