w<#
.Description
	This script to auto convert all doc/docx files to pdf
.Example
	.\DocToPdf.ps1 -dir D:\path\to\docs -out D:\path\to\output -opt 0 -update $false
.Parameter
 dir - path to docx folder
 out - (optional) specify output diectory for pdf. Default is the $dir
 opt - (optional) choose quality, 1 - export for preview, smaller size, 0 - export for print, large file. Default 0
 update - (optional) $true - update fields, $false - don't update. Default is $true
#>

param (
	[Parameter (Mandatory=$true, Position=1)][string]$dir,
	[Parameter (Position=2)][string]$out = $dir.trim('\'),
	[Parameter (Position=3)][int]$opt = 0,
	[Parameter (Position=4)][boolean]$update = $true	
)

$path = $dir.trim('\')

$wd = New-Object -ComObject Word.Application
Get-ChildItem -Path $path -Include *.doc, *.docx -Recurse |
    ForEach-Object {
        $doc = $wd.Documents.Open($_.Fullname)
        $pdf = $out + '\' + $_.BaseName -replace $_.Extension, '.pdf'
		
		if ($update -eq $true) {
			$doc.Fields.Update() | Out-Null
			foreach ($Section in $doc.Sections)
			{
				# Update Header
				$Header = $Section.Headers.Item(1)
				$Header.Range.Fields.Update()
	
				# Update Footer
				$Footer = $Section.Footers.Item(1)
				$Footer.Range.Fields.Update()
			}
		}
		
		$doc.ExportAsFixedFormat($pdf,17,$false,$opt,0,0,$false,$false,1,$false,$false,$true)
        $doc.Close()
    }
$wd.Quit()

#expression.ExportAsFixedFormat(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, Range, From, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr)
# https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb256835(v=office.12)
