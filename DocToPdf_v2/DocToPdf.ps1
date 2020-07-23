<#
.Description
	This script to auto convert all doc/docx files to pdf
.Example
	.\DocToPdf.ps1 -dir D:\path\to\docs -opt 0 -update $false
.Parameter
 dir - path to doc/docx folder. pdfs are written to the same folder 
 opt - (optional) choose quality, 1 - export for web, smaller size, 0 - export for print, large file. Default 0
 update - (optional) $true - update fields and TOC, $false - don't update. Default is $true
#>

param (
	[Parameter (Position=1)][string]$dir = $(get-location),
	[Parameter (Position=2)][int]$opt = 0,
	[Parameter (Position=3)][boolean]$update = $true	
)

$path = $dir.trim('\')
$exclude = "old|_old|source|_source"

$wd = New-Object -ComObject Word.Application
Get-ChildItem -Path $path -Include *.doc, *.docx -Recurse -File | Where-Object {$_.PSParentPath -notmatch $exclude} |
    ForEach-Object {
        $doc = $wd.Documents.Open($_.Fullname)
        <#$pdf = $out + '\' + $_.Name -replace $_.Extension, '.pdf'#>
		$pdf = $_.Fullname -replace $_.Extension, '.pdf'
		
		Write-Host "Processing $_"
		
		if ($update -eq $true) {
			$doc.Fields.Update() | Out-Null
			foreach ($Section in $doc.Sections)
			{
				# Update Header
				$Header = $Section.Headers.Item(1)
				$Header.Range.Fields.Update() | Out-Null
	
				# Update Footer
				$Footer = $Section.Footers.Item(1)
				$Footer.Range.Fields.Update() | Out-Null
			}
			try {
				$doc.TablesOfContents(1).Update()
			}
			catch [System.Runtime.InteropServices.COMException] {
				Write-Host "WARNING TOC not found" -ForegroundColor Yellow
			}
			$range = $doc.content
			$wordFound = $range.Find.Execute("Ошибка!")
			if ($wordFound) { 
				Write-Host "ERROR   Updated field error" -ForegroundColor Red
			}
		}
		
		$doc.ExportAsFixedFormat($pdf,17,$false,$opt,0,0,$false,$false,1,$false,$false,$true)
        $doc.Close()
    }
$wd.Quit()

#expression.ExportAsFixedFormat(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, Range, From, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr)
# https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb256835(v=office.12)