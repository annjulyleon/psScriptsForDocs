<#
.Description
	This script adds and updates .vsd,.vsdx builtin properties from .xml config
.Example
	.\UpdateVsdProps_v1.ps1 -dir D:\projects\myproject\ -conf D:\projects\myproject\config.xml
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
$visio = New-Object -ComObject Visio.Application
$visio.Visible = $false

#example xml config (should contain vsdProperties node):
#<?xml version="1.0"?>
#<configuration>  
#  <vsdProperties>
#   <add key="Company" value="This is Company"/>
#	<add key="Category" value="This is Category"/>
#	<add key="Title" value="This is title"/>
#	<add key="Subject" value="This is subject"/>
#	<add key="Keywords" value="This is keywords"/>
#	<add key="Description" value="This is description"/>
#	<add key="Manager" value="This is manAGER"/>
#  </vsdProperties>
#</configuration>

if(Test-Path $conf) {
    Try {
        #Load config customProperties        
		$global:vsdProperties = @{}
        $config = [xml](get-content $conf)       
		
		foreach ($addNode in $config.configuration.vsdProperties.add) {
            if ($addNode.Value.Contains(‘;’)) {
                # Array case
                $value = $addNode.Value.Split(‘;’)
                    for ($i = 0; $i -lt $value.length; $i++) { 
                        $value[$i] = $value[$i].Trim() 
                    }
            }
            else {
                # Scalar case
                $value = $addNode.Value
            }
        $global:vsdProperties[$addNode.Key] = $value
        }
    }
    Catch [system.exception]{
    }
}

Get-ChildItem -Path $path -Include *.vsd, *.vsdx -Recurse -File | Where-Object {$_.PSParentPath -notmatch $exclude} | ForEach-Object {
        $doc = $visio.Documents.Open($_.Fullname)
		
		ForEach($vsdProperty in $vsdProperties.GetEnumerator())
		{
			$doc.($vsdProperty.Name) = $vsdProperty.Value
			Write-Host "Property " $vsdProperty.Name "set to value " $vsdProperty.Value
		}		
		$doc.save()
		$doc.close()
    }
	
if ($visio) 
	{
        $visio.Quit()
    }