' Funcion que abre varios o un libro de excel
' Parametros: Ruta(s) del libro(s)
' Retorna: True si existio un error y False si se ejecuta correctamente

' Protege el codigo. Permite retornar True si existio un error y False si se ejecuta correctamente
On Error Resume Next 

Dim Args

' Argumentos que se pasan al programa al momento de ejecucion
Set Args = WScript.Arguments

Set excelApp = GetObject(,"Excel.Application") ' Se crea la instancia global de excel

For Each Ruta In Args
    excelApp.Workbooks.Open Ruta
Next

' Retorna True si existio un error y False si se ejecuta correctamente
If Err.Number <> 0 Then
    WScript.Echo "Error en initExcels: " & Err.Description
    WScript.Quit True
Else
    WScript.Quit False
End If
