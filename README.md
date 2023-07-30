Tool to upload/process your excel spreadsheet to json file.
Can be parsed from a CSV and XLSX file, and XLSX file stored on your One-Drive,
or can be sent using a Excel Add-In

### [adding the Add-In to Excel](#add-in)
disclaimer:
> this has only been tested on my windows computer\
> and may be different another windows machine and will\
> definitely be different for a linux or mac machine

1. Open "Registry Editor".
1. At the top, paste `Computer\HKEY_USERS\S-1-5-21-206511805-761490701-4004513812-1001\SOFTWARE\Microsoft\Office\16.0\WEF\Developer`.
1. Click `Edit > New > String Value` and give it a name.
1. Double click on the name and set the value to the absolute path of the `ExcelAddIn/yoeman/manifest.xml`
1. For example: `C:\Users\{username}\Desktop\excelUpload\ExcelAddIn\yoeman/manifest.xml`
1. Next time you open Excel at the top there will be a "developer Add-In" you can use.

# [Requirements](#requirements)
### npm libraries
1. install using any of the following:
1. `./installDependencies.bat`,
1. `./installDependencies.sh`,
1. `npm install`

### Microsoft Graph
1. Register an app by following [these steps](https://learn.microsoft.com/en-us/graph/auth-register-app-v2).
1. Set your scope to at least `User.Read offline_access Files.ReadWrite`.
1. Save your client id and client secret as a `ClientId.txt` and `ClientSecret.txt` file respectively inside a `graphTokens` folder
1. Run the `serverStart` script and go to `http://localhost` and click the "Authenticate with Microsoft graph." button.

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