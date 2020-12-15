' Funcion que elimina la instancia de excel para que no quede en segundo plano ante posible detenimiento
' del programa
' Retorna: True si existio un error y False si se ejecuta correctamente

' Protege el codigo. Permite retornar True si existio un error y False si se ejecuta correctamente
On Error Resume Next 

Set excelApp = GetObject(,"Excel.Application")
Set inventorApp = GetObject(,"Inventor.Application")

excelApp.Quit

' Retorna True si existio un error y False si se ejecuta correctamente
If Err.Number <> 0 Then
    WScript.Echo "Error en closeMainInstances: " & Err.Description
    WScript.Quit True
Else
    WScript.Quit False
End If