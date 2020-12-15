' Funcion que cierra un libro o libros abiertos segun su nombre
' Parametros: Nombre(s) del libro(s) a cerrar
' Retorna: True si existio un error y False si se ejecuta correctamente

' Protege el codigo. Permite retornar True si existio un error y False si se ejecuta correctamente
On Error Resume Next 

Dim Args

' Argumentos que se pasan al programa al momento de ejecucion
Set Args = WScript.Arguments

Set excelApp = GetObject(,"Excel.Application") ' Se crea la instancia global de excel

For Each Nombre In Args
    excelApp.Workbooks.Close Nombre
Next

' Retorna True si existio un error y False si se ejecuta correctamente
If Err.Number <> 0 Then
    WScript.Echo "Error en closeExcels: " & Err.Description
    WScript.Quit True
Else
    WScript.Quit False
End If
