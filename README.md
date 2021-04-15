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


# Building

Clone a copy of the repository

`git clone https://github.com/MarcinMichnik-HiQ/Frends.Office.git`

Rebuild the project

`dotnet build --configuration Release`

Create a NuGet package

`dotnet pack Frends.Office.nuspec`
