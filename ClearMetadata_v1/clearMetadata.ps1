<#
.Description
	This script removes all meta information drom all doc, docx files in the folder (including custom properties)
.Example
	.\clearMetadata.ps1 -path D:\projects\myproject\
.PARAMETER
	path - path to docx folder. Default is current folder
#>

param (
	[Parameter (Position=1)][string]$path = $(get-location),
	[Parameter (Position=2)][datetime]$date = $(get-date)
)

Add-Type -AssemblyName Microsoft.Office.Interop.Word
$WdRemoveDocType = "Microsoft.Office.Interop.Word.WdRemoveDocInfoType" -as [type] 
$files = Get-ChildItem -Path $path -include *.doc, *.docx -recurse 
$application = New-Object -ComObject word.application 

foreach($file in $files) 
{ 
    $file.CreationTime = $date
	$file.LastAccessTime = $date
	$file.LastWriteTime = $date
	$documents = $application.Documents.Open($file.fullname) 
    "Removing document information from $file"    
    $documents.RemoveDocumentInformation($WdRemoveDocType::wdRDIAll)
    $documents.Save() 
    $application.documents.close() 
} 
$application.Quit()