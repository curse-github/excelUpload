Tool to upload/process your excel spreadsheet to json file.
Can be parsed from a CSV and XLSX file, and XLSX file stored on your One-Drive,
or can be sent using a Excel Add-In

### adding the Add-In to Excel
disclaimer\
> this has only been tested on my windows computer \
> and may be different another windows machine and will\
> definitely be different for a linux or mac machine\

&nbsp;&nbsp;&nbsp; Open "Registry Editor".\
&nbsp;&nbsp;&nbsp; At the top, paste `Computer\HKEY_USERS\S-1-5-21-206511805-761490701-4004513812-1001\SOFTWARE\Microsoft\Office\16.0\WEF\Developer`.\
&nbsp;&nbsp;&nbsp; Click `Edit > New > String Value` and give it a name.\
&nbsp;&nbsp;&nbsp; Double click on the name and set the value to the absolute path of the `ExcelAddIn/yoeman/manifest.xml`\
&nbsp;&nbsp;&nbsp; For example: `C:\Users\{username}\Desktop\excelUpload\ExcelAddIn\yoeman/manifest.xml`\
&nbsp;&nbsp;&nbsp; Next time you open Excel at the top there will be a "developer Add-In" you can use.

# [Requirements](#requirements)
### &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;npm libraries
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;install using any of the following:\
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`./installDependencies.bat`,\
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`./installDependencies.sh`,\
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`npm install`

### &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Microsoft Graph
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Register an app by following [these steps](https://learn.microsoft.com/en-us/graph/auth-register-app-v2).\
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set your scope to at least `User.Read offline_access Files.ReadWrite`.\
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Save your client id and client secret as a `ClientId.txt` and `ClientSecret.txt` file respectively inside a `graphTokens` folder\
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Run the `serverStart` script and go to `http://localhost` and click the "Authenticate with Microsoft graph." button.

# [Scripts](#scripts)

|Name          |Description                                                                                                                                                         |
|:-------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|csvParser     |Parses a CSV file to JSON                                                                                                                                           |
|serverStart   |Starts server used to [authenticate with Microsoft Graph](https://learn.microsoft.com/en-us/graph/auth-v2-user?tabs=http)                                           |
|              |and receives data from the excel Add-in made using [yeoman](https://www.npmjs.com/package/yo) and [generator-office](https://www.npmjs.com/package/generator-office)|
|OneDriveUpload|parses XLSX from your One Drive using [Microsoft Graph](https://learn.microsoft.com/en-us/graph/overview)                                                           |
|xlsxParser    |parses an XLSX file from the home directory of this project using [excelJs](https://www.npmjs.com/package/exceljs)                                                  |

# [Config](#config)

|Name     |Description                                              |Default value  |
|:--------|:--------------------------------------------------------|:--------------|
|filename |filename used to search for XLSX and CSV files           |"excelUpload"  |
|         |and name for output JSON file                            |               |
|sheetName|the name of the sheet inside the XLSX file               |"Sheet1"       |
|left     |left column of the search bounds, used for OneDriveUpload|"A"            |
|right    |right column of the search bounds                        |"E"            |