# Powershell scripts for docs

[RU](README_RU.md)

[UpdateDocxProps.ps1](#updatedocxpropsps1)  
[DocToPdf.ps1](#doctopdfps1)

## UpdateDocxProps.ps1

Add or update custom properties .docx for all documents in folder. Script updates all fields, headers and footers on save. Specify any number of properties via .xml configuration file.

Example usage:
```bash
.\UpdateDocxProps.ps1 -dir D:\path\to\docs -conf D:\path\to\config\UpdateDocxPropsConfig.xml
```
Example config:

```xml
<?xml version="1.0"?>
<configuration>
  <appSettings>
<!--Vars -->
    <add key="NameOfProperty1" value="ValueOfProperty1"/>
	<add key="NameOfProperty2" value="ValueOfProperty2"/>
  </appSettings>
</configuration>
```
Source: 
- [Powershell: Everything you wanted to know about hashtables](https://powershellexplained.com/2016-11-06-powershell-hashtable-everything-you-wanted-to-know-about/)
- [How can I introduce a config file to Powershell scripts?](https://stackoverflow.com/a/13698982)
- [How to change custom properties for many Word documents](https://stackoverflow.com/a/35920682)
- [Powershell Update Fields in Header and Footer in Word](https://stackoverflow.com/questions/24887905/powershell-update-fields-in-header-and-footer-in-word)

## DocToPdf.ps1

Convert .docx and .doc to pdf + update fields. Output pdf are placed in the same directory.

Usage:
```bash
.\DocToPdf.ps1 -dir D:\path\to\docs -opt 1
```
`-dir` path to folder with doc/docx  
`-opt` (optional) choose quality, 1 - export for print, smaller size, 0 - export for print, large file. Default 0

Source: 
- [powershell script convert doc to pdf](https://social.technet.microsoft.com/Forums/ie/en-US/445b2429-e33c-4ce0-9d64-dd31422571bf/powershell-script-convert-doc-to-pdf?forum=winserverpowershell)
- [Document.ExportAsFixedFormat Method](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb256835(v=office.12))
