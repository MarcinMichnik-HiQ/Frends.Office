# Frends.Office

FRENDS Task for writing Excel files in .xsls format.

# Tasks

## WriteExcel

Reads csv string and converts it to an excel file.

### Properties

| Property | Type | Description | Example |
| -------- | -------- | -------- | -------- |
| CsvString | `string` | Input csv string. | `one;two;three\r\n1;2;3` |
| Delimiter | `char` | Determines what character will be used for splitting based on cell in csv. | `;` |
| LineDelimiter | `char` | Determines what string will be used for splitting lines. | "\r\n" |
| IncludeHeaders | `bool` | If input csv includes column names (headers). | true |
| Path | `string` | Full path of the target file to be written. File format should be .xlsx. | "c:\temp\file.xlsx" |

### Returns

Boolean - true if successful.

## ExcelToSharepoint

Finds an excel file at given path and sends it to sharepoint via Microsoft Graph API.

### Properties

| Property | Type | Description | Example |
| -------- | -------- | -------- | -------- |
| path | `string` | Path to a local file. | `c:\temp\file.xlsx` |
| fileName | `string` | The full name and extension of the file (same as in path) | `file.xlsx` |
| clientID | `string` | Azure Active Directory Site ID | `1ce3f5e1-fc04-3f24-2c0e-v76d5b44b13c` |
| tenantID | `string` | Azure Active Directory tenant id | `3d426023-5x12-4s11-afae-159b1865eabc` |
| clientSecret | `string` | Azure Active Directory client secret password | `_Sgx6Jdi2NC1N27Z4_plRm55L-DeCWJ.yq` |
| siteID | `string` | Azure AD Site ID | `name.sharepoint.com,f7b1c426-4x3c-4a7e-2129-296ed8449b49` |
| driveID | `string` | Azure AD Drive ID | `b!JsSx1zxNfkeRKSlc2ESbSfp6TF09EspGo2ERaFyAbykdDwKUa4CuRaPGyaagjGIN	` |
| targetFolderName | `string` | Target folder path | `General/Folder/` |

### Returns

Boolean - true if successful.

# Building

Clone a copy of the repository

`git clone https://github.com/MarcinMichnik-HiQ/Frends.Office.git`

Rebuild the project

`dotnet build --configuration Release`

Create a NuGet package

`dotnet pack Frends.Office.nuspec`

# Usage in Frends
## WriteExcel
<img width="306" alt="2021-04-15_12h28_01" src="https://user-images.githubusercontent.com/81616998/114855260-09f4a900-9de6-11eb-88cf-5adb871ba7dd.png">
## ExcelToSharepoint
<img width="217" alt="2021-04-26_17h13_23" src="https://user-images.githubusercontent.com/81616998/116106792-b66e3f00-a6b2-11eb-95b9-cf2af55b616b.png">


