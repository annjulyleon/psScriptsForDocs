# PowerShell Snippets

## Remove all files with extension

Removes file recursively.

````powershell
Get-ChildItem -Recurse *.pdf | foreach { Remove-Item -Path $_.FullName }
```
````

## Rename files

```powershell
Get-ChildItem *.docx -Filter 'RegExToFilterFiles*' | ForEach { Rename-Item $_ -NewName "NewName" }
```

Example for renaming multiple files to add folder name.

```powershell
$path = Split-Path -Path (Get-Location) -Leaf
$path = $path -replace 'П\.', ''
$path = $path -replace '\.РД',''
write-host $path

Get-ChildItem *.docx -Filter '00-*' | ForEach { Rename-Item $_ -NewName "00_${path}_Титул_v5-0.docx" }
```

## Count docx pages

```powershell
$word = New-Object -ComObject Word.Application 
$word.Visible = $false
...
$p2Pages = $word.ActiveDocument.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticPages)
```

Example code to select the document, count pages and write pages number to custom property.

```powershell
param (
	[Parameter (Position=1)][string]$dir = $(get-location)
)
$path = $dir.trim('\')

$word = New-Object -ComObject Word.Application 
$word.Visible = $false 

Get-ChildItem *.docx -Filter '*RegExp-To-Filter-byName*' | 
    ForEach-Object {  
        $doc = $word.Documents.Open($_.Fullname)
        $p2Pages = $word.ActiveDocument.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticPages)
        $doc.Close(0)
    }

function AddOrUpdateCustomProperty ($CustomPropertyName, $CustomPropertyValue, $DocumentToChange)
{
    $customProperties = $DocumentToChange.CustomDocumentProperties
    
    $binding = "System.Reflection.BindingFlags" -as [type]
    [array]$arrayArgs = $CustomPropertyName,$false, 4, $CustomPropertyValue
    Try 
    {
       [System.__ComObject].InvokeMember("add", $binding::InvokeMethod,$null,$customProperties,$arrayArgs) | out-null
    } 
    Catch [system.exception] 
    {
        $propertyObject = [System.__ComObject].InvokeMember("Item", $binding::GetProperty, $null, $customProperties, $CustomPropertyName)
        [System.__ComObject].InvokeMember("Delete", $binding::InvokeMethod, $null, $propertyObject, $null)
        [System.__ComObject].InvokeMember("add", $binding::InvokeMethod, $null, $customProperties, $arrayArgs) | Out-Null
    }
    Write-Host -ForegroundColor Green "Success! Custom Property:" $CustomPropertyName "set to value:" $CustomPropertyValue
}

Get-ChildItem *.docx -Filter 'RegExFilter for docs*' | 
    ForEach-Object {  
        $doc = $word.Documents.Open($_.Fullname)
        
        AddOrUpdateCustomProperty "p2Pages" $p2Pages $doc        
        $doc.Saved = $false
        $doc.save()
        $doc.Close(0)
    }

$word.Quit()
```

