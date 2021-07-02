@echo OFF
REM chcp 1251

Set CurrentDate=%Date%
Set CurrentTime=%Time: =0%
Set CurrentDateTime=%CurrentDate:~6,4%_%CurrentDate:~3,2%_%CurrentDate:~0,2%_%CurrentTime:~0,2%_%CurrentTime:~3,2%_%CurrentTime:~6,2%
 
Powershell.exe -noprofile -executionpolicy bypass -File UpdateDocxProps_v3.ps1 > %CurrentDateTime%_docprops.txt
Powershell.exe -noprofile -executionpolicy bypass -File UpdateVsdProps_v1.ps1 > %CurrentDateTime%_vsdprops.txt

