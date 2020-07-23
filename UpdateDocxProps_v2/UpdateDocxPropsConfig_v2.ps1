<#
.Description
	This script adds and updates .docx custom properties from .xml config
.Example
	.\UpdateDocxPropsConfig_v2.ps1 -dir D:\projects\myproject\ -conf D:\projects\myproject\config.xml
.PARAMETER
 dir - path to docx folder. Default is current folder
 conf - path to xml with configuration. Default is config.xml in current folder
#>

param (
	[Parameter (Position=1)][string]$dir = $(get-location),
	[Parameter (Position=2)][string]$conf = $dir + '\config.xml'
)

$path = $dir.trim('\')
$exclude = "old|_old|source|_source"

write-host "Working directory: " $dir 
write-host "Conig file: " $conf

# Read config file and load values
# Example XML:
# <configuration>
#   <startup>
#   </startup>
#   <appSettings>
# <!--Vars -->
#     <add key="projectCode" value="SomeValue"/>
#   </appSettings>
# </configuration>

if(Test-Path $conf) {
    Try {
        #Load config appsettings
        $global:appSettings = @{}
        $config = [xml](get-content $conf)
        foreach ($addNode in $config.configuration.appsettings.add) {
            if ($addNode.Value.Contains(‘,’)) {
                # Array case
                $value = $addNode.Value.Split(‘,’)
                    for ($i = 0; $i -lt $value.length; $i++) { 
                        $value[$i] = $value[$i].Trim() 
                    }
            }
            else {
                # Scalar case
                $value = $addNode.Value
            }
        $global:appSettings[$addNode.Key] = $value
        }
    }
    Catch [system.exception]{
    }
}

write-host "Working directory: " $path
$application = New-Object -ComObject word.application
$application.Visible = $false

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

Get-ChildItem -Path $path -Include *.doc, *.docx -Recurse -File | Where-Object {$_.PSParentPath -notmatch $exclude} |
    ForEach-Object {
        $doc = $application.Documents.Open($_.Fullname)      
		
		Write-Host "Processing $_"
		
		ForEach($property in $appSettings.GetEnumerator())
		{
			AddOrUpdateCustomProperty $($property.Name) $($property.Value) $doc
		}

		Write-Host "Updating document fields."
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

		Write-Host "Saving document."	
		
        $doc.Saved = $false
		$doc.save()
		$doc.close()
    }
$application.Quit()
$application = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()