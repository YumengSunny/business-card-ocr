
*****************
Deployement Steps
*****************

1) Go to azure account. Click on Create new resource. Select Computer Vision and complete the process by filling simple form. You will be provided key to access this service.
   Copy that key and replace with value of variable subscriptionKey in CardReaderController.cs.

2) Go to azzure account. Click on Create new resource. Select Text Analytics and complete the process by filling simple form. You will be provided key to access this service.
   Copy that key and replace with value of variable cognitiveTextApiKey in CardReaderController.cs and GoogleReaderController.cs.  
 
3) Keep json file of google vision's api to D:\home\myapp\apikeys.json and give this path in Settings >Application settings> Application settings of deployeed app.
 
4) Make sure that python is installed with open-cv and argparse packages. Replace python's path with value of variable pythonpath in CardReaderController.cs and GoogleReaderController.cs.  Keep imagesegmentation.py on same path. 

5) Create folder D:\\home\\images\\google\\ with full permission. 

6) Create folder D:\\home\\images\\azure\\ with full permission. 
   
7) Follow the steps given in link to deploy the web application in azure.
https://docs.microsoft.com/en-us/azure/app-service/app-service-web-get-started-dotnet

8) For SQL intergration, you will have to create SQL server named 'sqlserverocr' with username 'developer1' and your password. 

9) Now, please go to the server from all resources and click on 'Firewalls and virtual networks Show firewall settings'.

10) Add your PC Ip address to connect to the SQL server.

11) Open SQL server management studio and connect to Azure SQL server.

12) Import and Export Database: Follow Step #10 in my resource group's SQL server and connect that SQL server into your SQL Server Management Studio.

13) Once you connect to both new and old server, please do Import and Export of user_data database. 

14) Note: You can find the connection string to connect to database from Azure portal by clicking on your Database. 

