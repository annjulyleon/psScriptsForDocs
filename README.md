# Powershell scripts for docs

[RU](README_RU.md#powershell-scripts-for-docs)

[UpdateDocxProps.ps1](#updatedocxpropsps1)  
[DocToPdf.ps1](#doctopdfps1)  
[FindAndReplace.ps1](#findandreplaceps1)

## UpdateDocxProps.ps1

Add or update custom properties .docx for all documents in folder. Script updates all fields, headers and footers on save. Specify any number of properties via .xml configuration file.

**v2**: 

* code refactoring
* exclude folders (`exclude` variable), default is `old,_old,source,_source`
* example bat-file with logging added

Example usage:
```bash
.\UpdateDocxProps.ps1 -dir D:\path\to\docs -conf D:\path\to\config\UpdateDocxPropsConfig.xml
```

Use with .bat:
1. copy script files to doc/docx folder
2. change config.xml, add properties
3. launch updateProps.bat with administrative rights (to overwrite ps restriction). Uncomment "chcp 1251" string if using with RU language
4. check log file

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

Convert .docx and .doc to pdf + update fields (optional).

**v2**:

* remove `-out` parameter. Saves the pdf to the original docx folder
* example bat-file with logging added
* vsd/vsdx to pdf script added (separate file)
* error checking added for broken links (see below to fix script for your language, Russian is default)
* update TOC
* exclude folders с `old, _old, source, _source` 

Usage **v1**:

```bash
.\DocToPdf.ps1 -dir D:\path\to\docs -out D:\path\to\output -opt 0 -update $false
```
`-dir` - path to docx folder  
`-out` - (optional) specify output directory for pdf. Default is the `$dir`  
`-opt` - (optional) choose quality, `1` - export for preview, smaller size, `0` - export for print, large file. Default 0  
`-update` - (optional) `$true` - update fields, `$false` - don't update. Default is `$true`  

Usage **v2**:

```
.\DocToPdf.ps1 -dir D:\path\to\docs -opt 0 -update $false
```

`-dir` - (optional) path to docx folder. Default is current folder
`-opt` - (optional) - 0/1 - choose quality, 1 - export for preview, smaller size, 0 - export for print, large file. Default is 0  
`-update` - (optional) true/false - `$true` - update fields, `$false` - don't update. Default is `$true`  

Included example `.bat` file write errors and warnings to the log-file:

- WARNING - if there is no TOC in document
- ERROR - if key error word is found (for ex. broken link is found)

You need to change key word for the broken link error text for your language:

```powershell
$wordFound = $range.Find.Execute("Ошибка!")
```

Source: 

- [powershell script convert doc to pdf](https://social.technet.microsoft.com/Forums/ie/en-US/445b2429-e33c-4ce0-9d64-dd31422571bf/powershell-script-convert-doc-to-pdf?forum=winserverpowershell)
- [Document.ExportAsFixedFormat Method](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb256835(v=office.12))


## FindAndReplace.ps1

Find and replace strings for multiple .docx files

Usage:
```bach
.\FindAndReplace.ps1 D:\path\to\folder 'text to find' 'text to replace'
```
`-dir` - directory with docx files  
`-find` - string to find  
`-replace` - string to replace  

Source and usefull links:
- [Replacing many Words in a .docx File with Powershell](https://stackoverflow.com/questions/40101846/replacing-many-words-in-a-docx-file-with-powershell)
- [PowerShell script to Find and Replace in Word Document, including Header, Footer and TextBoxes within
](https://codereview.stackexchange.com/questions/174455/powershell-script-to-find-and-replace-in-word-document-including-header-footer)
