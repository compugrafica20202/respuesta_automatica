' Funcion que crea una instancia global de la aplicacion de excel y abre varios o un libro
' Retorna: True si existio un error y False si se ejecuta correctamente

' Protege el codigo. Permite retornar True si existio un error y False si se ejecuta correctamente
On Error Resume Next 

Dim excelApp
Dim inventorApp

Set excelApp = CreateObject("Excel.Application") ' Se crea la instancia global de excel
excelApp.Visible = False 'Importante sea False para mejorar el rendimiento
excelApp.DisplayAlerts = False 'Importante para que no hayan interrupciones
excelApp.Workbooks.Add 

Set inventorApp = CreateObject("Inventor.Application")
inventorApp.Visible = False

' Retorna True si existio un error y False si se ejecuta correctamente
If Err.Number <> 0 Then
    WScript.Echo "Error en initExcels: " & Err.Description
    WScript.Quit True
Else
    WScript.Quit False
End If