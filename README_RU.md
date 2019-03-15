# Powershell scripts for docs

[UpdateDocxProps.ps1](#updatedocxpropsps1)  
[DocToPdf.ps1](#doctopdfps1)  
[FindAndReplace.ps1](#findandreplaceps1)

## UpdateDocxProps.ps1

Скрипт PowerShell для добавления и обновления свойств в документах .docx. При сохранении обновляет поля, в том числе в колонтитулах. Свойства берутся из конфигурационного файла .xml.

Запуск:

```bash
.\UpdateDocxProps.ps1 -dir D:\path\to\docs -conf D:\path\to\config\UpdateDocxPropsConfig.xml
```

Пример конфига:

```xml
<?xml version="1.0"?>
<configuration>
  <appSettings>
<!--Vars -->
    <add key="NameOfProperty1" value="ValueOfProperty1"/>
	<add key="NameOfProperty2" value="ValueOfProperty2"/>
  </appSettings>
</configuration>
```

Источники: 

- [Powershell: Everything you wanted to know about hashtables](https://powershellexplained.com/2016-11-06-powershell-hashtable-everything-you-wanted-to-know-about/)
- [How can I introduce a config file to Powershell scripts?](https://stackoverflow.com/a/13698982)
- [How to change custom properties for many Word documents](https://stackoverflow.com/a/35920682)
- [Powershell Update Fields in Header and Footer in Word](https://stackoverflow.com/questions/24887905/powershell-update-fields-in-header-and-footer-in-word)

# DocToPdf.ps1

Конвертирует документы doc/docx в pdf. Обновляет поля (опционально), можно настраивать качество pdf (для просмотра или для печати)

Github: <https://github.com/annjulyleon/psScriptsForDocs>

Использование:
```bash
.\DocToPdf.ps1 -dir D:\path\to\docs -out D:\path\to\output -opt 0 -update $false
```
`-dir` - путь к папке с docx   
`-out` - (необязательно) папка для выходных pdf, по умолчанию `$dir`  
`-opt` - (необязательно) качество файла pdf, 1 - экспорт для веба и предпросмотра, меньший файл, 0 - экспорт для печати, большой файл. По умолчанию 0  
`-update` - (необязательно) `$true` - обновить поля документа перед сохранением, `$false` - не обновлять. По умолчанию `$true`  

Источники: 
- [powershell script convert doc to pdf](https://social.technet.microsoft.com/Forums/ie/en-US/445b2429-e33c-4ce0-9d64-dd31422571bf/powershell-script-convert-doc-to-pdf?forum=winserverpowershell)
- [Document.ExportAsFixedFormat Method](https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb256835(v=office.12))

## FindAndReplace.ps1

 Поиск и замена текстовой строки для нескольких .docx файлов

Использование:
```bach
.\FindAndReplace.ps1 D:\path\to\folder 'text to find' 'text to replace'
```
`-dir` - путь к папке с docx  
`-find` - строка для поиска  
`-replace` - строка для замены  

Источники и полезные ссылки:
- [Replacing many Words in a .docx File with Powershell](https://stackoverflow.com/questions/40101846/replacing-many-words-in-a-docx-file-with-powershell)
- [PowerShell script to Find and Replace in Word Document, including Header, Footer and TextBoxes within
](https://codereview.stackexchange.com/questions/174455/powershell-script-to-find-and-replace-in-word-document-including-header-footer)
