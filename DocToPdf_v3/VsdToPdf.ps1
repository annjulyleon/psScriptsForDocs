param (
	[Parameter (Position=1)][string]$dir = $(get-location)
)

$exclude = "old|_old|source|_source"
$path = $dir.trim('\')
$visio = New-Object -ComObject Visio.Application
$visio.Visible = $false

Get-ChildItem -Path $path -Include *.vsd, *.vsdx -Recurse -File | Where-Object {$_.PSParentPath -notmatch $exclude} | ForEach-Object {
        $doc = $visio.Documents.Open($_.Fullname)
		$pdf = $_.Fullname -replace $_.Extension, '.pdf'
		
		Write-Host "Processing $_"		
		$doc.ExportAsFixedFormat(1,$pdf,1,0)
        $doc.Close()
    }
	
if ($visio) 
	{
        $visio.Quit()
    }