<#
.Description
	This script to auto convert all doc/docx files to pdf
.Example
	.\DocToPdf.ps1 -dir D:\path\to\docs -opt 0
.Parameter
 dir - path to docx folder
 opt - (optional) choose quality, 1 - export for print, smaller size, 0 - export for print, large file. Default 0
#>

param (
	[Parameter (Mandatory=$true, Position=1)][string]$dir,
	[Parameter (Position=2)][int]$opt = 0
)
$path = $dir

$wd = New-Object -ComObject Word.Application
Get-ChildItem -Path $path -Include *.doc, *.docx -Recurse |
    ForEach-Object {
        $doc = $wd.Documents.Open($_.Fullname)
        $pdf = $_.FullName -replace $_.Extension, '.pdf'
        $doc.ExportAsFixedFormat($pdf,17,$false,$opt,0,0,$false,$false,1,$false,$false,$true)
        $doc.Close()
    }
$wd.Quit()

#expression.ExportAsFixedFormat(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, Range, From, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr)
# https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb256835(v=office.12)