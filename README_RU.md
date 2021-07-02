# Powershell scripts for docs

[UpdateDocxProps.ps1](#updatedocxpropsps1)  
[DocToPdf.ps1](#doctopdfps1)  
[FindAndReplace.ps1](#findandreplaceps1)
[Clear metadata](#clearmetadata)

## UpdateDocxProps.ps1

Скрипт PowerShell для добавления и обновления свойств в документах .docx. При сохранении обновляет поля, в том числе в колонтитулах. Свойства берет из конфигурационного файла .xml.

### v3

* теперь скрипт умеет обновлять встроенные свойства в Visio файлах
* обновляет встроенные свойства в Word (Теги, Примечания, Тема, Руководитель...)
* скрипт обновляет поля внутри форм (shapes) и надписей
* изменен формат конфигурационного файла (секции для vsd и встроенных свойств Word)

Структура файлов:

* `config.xml` - конфигурационный файл. Теперь включает три секции: `customProperties` для кастомных свойств Word, `builtinProperties` - встроенные свойства Word, `vsdProperties` - встроенные Visio свойства. Пример:
  
    ```xml
    <?xml version="1.0"?>
    <configuration>
      <customProperties>    
        <add key="property1" value="new value for the property"/>	
      </customProperties>
      <builtinProperties>
        <add key="Title" value="This is Title property"/>
        <add key="Subject" value="This is Subject property"/>
        <add key="Keywords" value="some tag more tag"/>
        <add key="Comments" value="somecomment"/>
      </builtinProperties>
      <vsdProperties>
        <add key="Company" value="LLC COMPANY"/>
        <add key="Category" value="Category of the document"/>
        <add key="Title" value="Title of the document"/>
        <add key="Subject" value="Subject of the document"/>
        <add key="Keywords" value="Some tags"/>
        <add key="Description" value="Desc comment"/>
        <add key="Manager" value="Project Manager"/>
      </vsdProperties>
    </configuration>
    ```
  Если какие-то свойства или вся секциия не нужны, то просто удалите все свойства в секции.

```xml
    <?xml version="1.0"?>
    <configuration>
      <customProperties>    
        <add key="property1" value="new value for the property"/>	
      </customProperties>
      <builtinProperties>        
      </builtinProperties>
      <vsdProperties>        
      </vsdProperties>
    </configuration>
```
* `updateProps.bat` запускает скрипт для текущией и дочерних директорий (за исключением директорий, указанных в `exclude` срикпта. Если не нужно запускать скрипт для vsd файлов, просто закомментируйте эту строку в .bat:
    ```bat
    Powershell.exe -noprofile -executionpolicy bypass -File UpdateVsdProps_v1.ps1 > %CurrentDateTime%_vsdprops.txt
    ```
    
* `UpdateDocxProps_v3.ps1` - скрипт для обновления doc/docx свойств;

* `UpdateVsdProps_v1.ps1` - обновляет vsd свойства;

* два тестовых файла: `testvsdfile.vsd` и `teswordfile.docx`.

Измените конфигурационный файл и запустите скрипт с помощью `.bat`. Оба скрипта имеют одинаковые параметры:`-dir` (по умолчанию - текущая директория) и `-conf` (по умолчанию - файл `config,xml`в текущей директории).


### v2: 

* рефакторинг кода до новой версии PowerShell
* исключение папок (переменная exclude в коде), по умолчанию исключаются папки `old,_old,source,_source`
* пример bat-файла для запуска, логирование

Запуск скрипта:

```bash
.\UpdateDocxProps.ps1 -dir D:\path\to\docs -conf D:\path\to\config\UpdateDocxPropsConfig.xml
```

Использование:

1. скопировать файлы в директорию, в которой лежат файлы для изменения (doc, docx)
2. поменять значения свойств в config.xml на нужные
3. запустить файл updateProps.bat от администратора. Если используется русский язык, то раскомментировать строку `chcp 1251` (убрать REM)
4. дождаться пока исчезнет окно консоли
5. проверить лог-файл на ошибки

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

Конвертирует документы doc/docx в pdf. Обновляет поля (опционально), можно настраивать качество pdf (для просмотра или для печати).

### v3:

* теперь скрипт умеет обновлять перед выводом в pdf поля в фигурах и надписях

### v2:

* убран параметр out. Теперь всегда pdf сохраняется в папку, где был документ docx. Сделано, чтобы можно было обрабатывать сразу много папок
* добавлен bat для запуска с логированием. Если используется русский язык, то нужно раскомментировать строку `chcp 1251` (убрать REM) в bat-файле
* добавлен конвертер для vsd/vsdx
* добавлена проверка ошибок в файле перед выводом, проверяется текст "Ошибка!", который Word пишет для битых ссылок. Если ошибка найдена - выводится в лог
* добавлено обновление содержания перед выводом pdf
* исключаются папки с `old, _old, source, _source` в названии

Использование **v1**:
```bash
.\DocToPdf.ps1 -dir D:\path\to\docs -out D:\path\to\output -opt 0 -update $false
```
`-dir` - путь к папке с docx   
`-out` - (необязательно) папка для выходных pdf, по умолчанию `$dir`  
`-opt` - (необязательно) качество файла pdf, 1 - экспорт для веба и предпросмотра, меньший файл, 0 - экспорт для печати, большой файл. По умолчанию 0  
`-update` - (необязательно) `$true` - обновить поля документа перед сохранением, `$false` - не обновлять. По умолчанию `$true`  

Использование **v2**:

```
.\DocToPdf.ps1 -dir D:\path\to\docs -opt 0 -update $false
```

`-dir` - путь к папке с docx файлами. Если не указан, то выполняется для файлов в папке скрипта  
`-opt` - 0/1 - качество вывода на печать - 1 для веба (маленький файл), 0 на печать. По умолчанию, если не указан = 0  
`-update` - true/false - true - обновлять поля и содержание, false - не обновлять. По умолчанию и если не указано - true  

Скрипт пишет вывод для каждого документа (показывает полный путь) в  лог-файл с датой. Скрипт выполняет простейшую проверку документа и  выводит: 

- WARNING - если в документе нет содержания (нормальная ситуация для спецификаций, ведомостей и пр.)
- ERROR - если в документе обнаружен текст "Ошибка!". Чаще всего это означает, что найдена ошибка обновления поля или отсутствует какое-то  свойство в документе

Если требуется изменить ключевое слово для поиска ошибки (например, для другого языка), то изменить в скрипте строку (слово с скобках):

```powershell
$wordFound = $range.Find.Execute("Ошибка!")
```



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

## ClearMetadata

Небольшой скрипт, который удаляет все метаданные из файлов .docx и устанавливает все даты на текущую (Дата создания, Дата доступа, Дата сохранения). Включен .bat для запуска скрипта для текущей и дочерних директорий.

Для запуска вручную выполнить команду:

```
.\clearMetadata.ps1 -path D:\path\to\folder
```

`-path` - путь к папке с файлами, по умолчанию - текущая, в которой находится скрипт.