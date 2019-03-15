<#
.Description
	Simple find and replace for multiple docs
.Example
	.\FindAndReplace.ps1 D:\path\to\folder 'text to find' 'text to replace'
.Parameter
  dir - directory with docx files
  find - string to find
  replace - string to replace
#>
param (
	[Parameter (Mandatory=$true, Position=1)][string]$dir,
	[Parameter (Mandatory=$true, Position=2)][string]$find,
	[Parameter (Mandatory=$true, Position=3)][string]$replace
)

#$path = $dir 

$application = New-Object -ComObject word.application
$application.Visible = $false

ForEach($File in (GCI $dir|where {$_.extension -eq ".docx"}|Select -Expand FullName))
{
    Write-Host "Opening Document..." $File
    $document = $application.documents.open($File)
	$objSelection = $application.Selection
	
    $FindText = $find
	$ReplaceWith = $replace
    
	$MatchCase = $False 
    $MatchWholeWord = $False 
    $MatchWildcards = $False 
    $MatchSoundsLike = $False 
    $MatchAllWordForms = $False 
    $Forward = $True 
    $Wrap = $wdFindContinue 
    $Format = $False 
    $wdReplaceNone = 1     
    $wdFindContinue = 1 
	
    $objSelection.Find.Execute($FindText,$MatchCase,$MatchWholeWord,$MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,$Wrap,$Format,$ReplaceWith,2)
	
    $document.Saved = $false
    $document.save()
    $document.close()
}

$application.quit()
$application = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()

write-host "Done!"
write-host "Press any key to continue..."
[void][System.Console]::ReadKey($true)
