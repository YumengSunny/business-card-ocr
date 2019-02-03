*******************************************
Deployment Steps
*******************************************

1) Follow steps given in the link to deploy the web application from Visual Studio to azure.
https://docs.microsoft.com/en-us/azure/app-service/app-service-web-get-started-dotnet

2) From Azure account, click on Create new resource. Select Computer Vision and complete the process by filling simple form. You will be provided key to access this service. Copy that key and replace with value of variable subscriptionKey in CardReaderController.cs.

3) From Azure account, click on Create new resource. Select Text Analytics and complete the process by filling simple form. You will be provided key to access this service. Copy that key and replace with value of variable cognitiveTextApiKey in CardReaderController.cs and GoogleReaderController.cs.  
 
4) From Azure App setting Extention, please install latest version of Python. 

5) Replace python's path with value of variable pythonpath in CardReaderController.cs and GoogleReaderController.cs. 
Important Note: You will have to set environmental path from Kudu for Python, and only use the path which was used in above step. 

6) Keep imagesegmentation.py on same path. 

7) Now follow these steps in Kudu console plateform

-- Create folder D:\\home\\images\\google\\ with full permission.

-- Create folder D:\\home\\images\\azure\\ with full permission.

-- Keep json file of google vision's api to D:\home\myapp\apikeys.json and give this path in Settings > Application settings > Application settings of deployeed app.

-- Install Opencv package with version 3.2.0.8. Note: Do not install other version.

-- Install argparse package.

8) SQL Installation:

-- You will have to create SQL server named 'sqlserverocr' with username 'developer1' and your password.

-- Now, please go to the server from all resources and click on 'Firewalls and virtual networks Show firewall settings'.

-- Add your PC Ip address to connect to the SQL server.

-- Follow same path in my resource group's SQL server and connect the old SQL server into your SQL Server Management Studio.

-- Once you connect to both new and old server, please do Import and Export of user_data database. 

-- Note: You can find the connection string to connect to database from Azure portal by clicking on your Database. 

-- Use the connection string in the project.
