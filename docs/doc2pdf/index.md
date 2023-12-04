# Преобразование в PDF

## Оглавление

1. [Для ОС Windows](students/index.md)
2. [Универсальный способ](curators/index.md)

---

### Операционная система Windows

1. Создать файл doc2pdf.vbs (или [скачать](Doc2PDF.vbs)) и скопировать в него следующий код
```vb
'Convert .doc or .docx to .pdf files via Send To menu
Set fso = CreateObject("Scripting.FileSystemObject")
For i= 0 To WScript.Arguments.Count -1
   docPath = WScript.Arguments(i)
   docPath = fso.GetAbsolutePathName(docPath)
   If LCase(Right(docPath, 4)) = ".doc" Or LCase(Right(docPath, 5)) = ".docx" Then
      Set objWord = CreateObject("Word.Application")
      pdfPath = fso.GetParentFolderName(docPath) & "\" & _
	fso.GetBaseName(docpath) & ".pdf"
      objWord.Visible = False
      Set objDoc = objWord.documents.open(docPath)
      objDoc.saveas pdfPath, 17
      objDoc.Close
      objWord.Quit
   End If
Next
```
2. Откройте папку для вашего пользователя с пунктами меню "Отправить"
3. 


> **авторы:** 
>   - Ляш Олег Иванович, 
> 
>  *Условия использования: любое использование материалов запрещено без письменного разрешения авторов*