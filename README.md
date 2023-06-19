- [Introduction](#introduction)
  - [Export](#export)
  - [Import](#import)
  - [Properties](#properties)

# Introduction
This PowerShell Script exports AD values to a CSV file.
This file can be given to HR so they can change and update departments, managers or phone numbers.
To put the data back, you need to save the file in CSV (with a ";" delimiter) and UTF-8 encoding (should be stadard)

## Export
To export AD User data, just type:
```PowerShell
    .\ExportImportADUsersByCSV.ps1 -action export -OU "OU=MainSite,OU=Company"
```
The CSV-Filename is located in the script (`ADUserExport.csv`). This can be changed in the script itself.
The OU needs only be the "OU" part, the domain will be requested via `Get-ADDomain`.
## Import
The import file needs to be the same name and also an UTF-8 encoded CSV file with a ";" delimiter.
With the Parameter `-testOnly:$true` you can test, which data would be changed.
When the Parameter is set to `-testOnly:$false` the data will be written to AD via `Set-ADUser`.
```PowerShell
    .\ExportImportADUsersByCSV.ps1 -action import -OU "OU=MainSite,OU=Company" -testOnly:$true
```
## Properties
With this script it is possible to read also non standard Properties from AD and set them also via this script.
To get all possible Properties use `Get-ADUser <user> -Properties * | Get-Member`
In the Line:
```
$Properties = "Name,Title,Department,Description,Company,OfficePhone,Office,WWWHomePage,Manager,sAMAccountName"
```
You can change the Properies and add others. As the script relies on the `sAMAccountName` this column must be included.
Best put it to the end, so noone bothers to change that. :-)