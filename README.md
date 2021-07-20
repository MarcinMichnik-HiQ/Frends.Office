# Frends.Office

FRENDS Task for writing Excel files in .xsls format.

# Tasks

## WriteExcelFileTask

Reads csv string, converts it to a DataTable and creates an excel file.

### Properties

| Property | Type | Description | Example |
| -------- | -------- | -------- | -------- |
| StringInput | `string` | Input csv string | `"one;two;three\r\n1;2;3"` |
| CellDelimiter | `char` | Determines what character will be used for splitting based on cell in csv | `';'` |
| LineDelimiter | `string` | Determines what string will be used for splitting lines | `"\r\n"` |
| TargetPath | `string` | Full path of the target file to be written. File format should be .xlsx | `@"c:\temp\file.xlsx"` |

### Returns

JToken with 'message' and 'filePath' keys.

## WriteWordFileTask

Reads string and creates a word file.

### Properties

| Property | Type | Description | Example |
| -------- | -------- | -------- | -------- |
| StringInput | `string` | Input text string | `"Paragraph\r\nNewLine"` |
| LineDelimiter | `string` | Determines what string will be used for splitting lines | `"\r\n"` |
| TargetPath | `string` | Full path of the target file to be written. File format should be .docx | `@"c:\temp\file.docx"` |

### Returns

JToken with keys: message, savedTo, rows.

## ExportFileToSharepointTask

Finds a file at given path and sends it to sharepoint via Microsoft Graph API.

### Properties

#### FileExportInput

| Property | Type | Description | Example |
| -------- | -------- | -------- | -------- |
| SourceFilePath | `string` | Path to a local file | `"c:\temp\file.xlsx"` |
| targetFolderName | `string` | Target folder path | `"General/Folder/"` |

#### Authentication

| Property | Type | Description | Example |
| -------- | -------- | -------- | -------- |
| clientID | `string` | Azure Active Directory Site ID | `"1ce3f5e1-fc04-3f24-2c0e-v76d5b44b13c"` |
| clientSecret | `string` | Azure Active Directory client secret password | `"_Sgx6Jdi2NC1N27Z4_plRm55L-DeCWJ.yq"` |
| tenantID | `string` | Azure Active Directory tenant id | `"3d426023-5x12-4s11-afae-159b1865eabc"` |
| siteID | `string` | Azure Active Directory Site ID | `"f7b1c426-4x3c-4a7e-2129-296ed8449b49"` |

### Returns

JToken with keys: FileSize, Path, FileName, TargetFolderName, ClientID, TenantID, SiteID, DriveID, UploadUrl.

# Building

Clone a copy of the repository

`git clone https://github.com/MarcinMichnik-HiQ/Frends.Office.git`

Build the project

`dotnet build --configuration Release`

Create a NuGet package

`nuget pack Frends.Office.nuspec`

